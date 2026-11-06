'use strict';
const express = require('express'), path = require('path'), fs = require('fs'), crypto = require('crypto');
const Database = require('better-sqlite3'), bcrypt = require('bcryptjs'), jwt = require('jsonwebtoken'), cookieParser = require('cookie-parser'), ExcelJS = require('exceljs');
const { invoicePDF, reportPDF } = require('./pdf');

// ───────── storage ─────────
const DATA = process.env.DATA_DIR || path.join(__dirname, 'data');
const DBFILE = path.join(DATA, 'billease.db'), INV = path.join(DATA, 'invoices'), BK = path.join(DATA, 'backups'), PEND = path.join(DATA, 'restore_pending.db');
[DATA, INV, BK].forEach(d => fs.mkdirSync(d, { recursive: true }));
const logErr = (e, ctx = '') => { try { fs.appendFileSync(path.join(DATA, 'error.log'), `[${new Date().toISOString()}] ${ctx} ${(e && e.stack) || e}\n`); } catch {} };
process.on('uncaughtException', e => logErr(e, 'uncaught')); process.on('unhandledRejection', e => logErr(e, 'unhandled'));
if (fs.existsSync(PEND)) { // apply a restore requested before restart
  if (fs.existsSync(DBFILE)) fs.copyFileSync(DBFILE, path.join(BK, `pre-restore-${Date.now()}.db`));
  for (const x of ['', '-wal', '-shm']) try { fs.unlinkSync(DBFILE + x); } catch {}
  fs.renameSync(PEND, DBFILE);
}
const db = new Database(DBFILE); db.pragma('journal_mode=WAL'); db.pragma('foreign_keys=ON');
db.exec(`
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL COLLATE NOCASE, name TEXT NOT NULL, password_hash TEXT NOT NULL,
  role TEXT NOT NULL CHECK(role IN('owner','manager','employee')), permissions TEXT, active INTEGER NOT NULL DEFAULT 1, created_by INTEGER REFERENCES users(id),
  created_at TEXT NOT NULL DEFAULT(datetime('now')), last_login TEXT);
CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY, sku TEXT UNIQUE NOT NULL, name TEXT NOT NULL, price REAL NOT NULL CHECK(price>=0), cost REAL NOT NULL DEFAULT 0,
  tax_pct REAL NOT NULL DEFAULT 0, discount_pct REAL NOT NULL DEFAULT 0, stock REAL NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT(datetime('now')), updated_at TEXT NOT NULL DEFAULT(datetime('now')));
CREATE TABLE IF NOT EXISTS customers(id INTEGER PRIMARY KEY, name TEXT NOT NULL, phone TEXT, address TEXT, created_at TEXT NOT NULL DEFAULT(datetime('now')));
CREATE TABLE IF NOT EXISTS orders(id INTEGER PRIMARY KEY, invoice_no TEXT UNIQUE NOT NULL, client_ref TEXT UNIQUE NOT NULL, customer_id INTEGER REFERENCES customers(id), customer_name TEXT,
  user_id INTEGER NOT NULL REFERENCES users(id), subtotal REAL NOT NULL, discount_total REAL NOT NULL, tax_total REAL NOT NULL, total REAL NOT NULL, paid REAL NOT NULL,
  prev_outstanding REAL NOT NULL DEFAULT 0, payment_method TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'completed', note TEXT,
  created_at TEXT NOT NULL DEFAULT(datetime('now')), cancelled_at TEXT, cancelled_by INTEGER REFERENCES users(id), cancel_reason TEXT);
CREATE TABLE IF NOT EXISTS order_items(id INTEGER PRIMARY KEY, order_id INTEGER NOT NULL REFERENCES orders(id), product_id INTEGER NOT NULL REFERENCES products(id), sku TEXT, name TEXT NOT NULL,
  qty REAL NOT NULL, unit_price REAL NOT NULL, cost REAL NOT NULL DEFAULT 0, discount_pct REAL NOT NULL DEFAULT 0, discount_amt REAL NOT NULL DEFAULT 0,
  tax_pct REAL NOT NULL DEFAULT 0, tax_amt REAL NOT NULL DEFAULT 0, line_total REAL NOT NULL, is_return INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY, customer_id INTEGER NOT NULL REFERENCES customers(id), amount REAL NOT NULL CHECK(amount>0), method TEXT NOT NULL, note TEXT,
  user_id INTEGER NOT NULL REFERENCES users(id), created_at TEXT NOT NULL DEFAULT(datetime('now')));
CREATE TABLE IF NOT EXISTS stock_movements(id INTEGER PRIMARY KEY, product_id INTEGER NOT NULL REFERENCES products(id), delta REAL NOT NULL, reason TEXT, user_id INTEGER REFERENCES users(id), created_at TEXT NOT NULL DEFAULT(datetime('now')));
CREATE TABLE IF NOT EXISTS activity_logs(id INTEGER PRIMARY KEY, user_id INTEGER, username TEXT, action TEXT NOT NULL, entity TEXT, entity_id TEXT, detail TEXT, ip TEXT, created_at TEXT NOT NULL DEFAULT(datetime('now')));
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
CREATE INDEX IF NOT EXISTS ix_orders_created ON orders(created_at); CREATE INDEX IF NOT EXISTS ix_orders_user ON orders(user_id); CREATE INDEX IF NOT EXISTS ix_orders_cust ON orders(customer_id);
CREATE INDEX IF NOT EXISTS ix_items_order ON order_items(order_id); CREATE INDEX IF NOT EXISTS ix_pay_cust ON payments(customer_id); CREATE INDEX IF NOT EXISTS ix_logs_created ON activity_logs(created_at);`);

const getS = k => { const r = db.prepare('SELECT value FROM settings WHERE key=?').get(k); return r ? r.value : null; };
const setS = (k, v) => db.prepare('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value').run(k, String(v));
const DEFAULT_SET = { business_name: 'My Business', business_address: '', business_phone: '', receipt_footer: 'Thank you!', route_name: '', tz_offset: '330', low_stock: '10', logo: '' };
for (const [k, v] of Object.entries(DEFAULT_SET)) if (getS(k) === null) setS(k, v);
const pubSettings = () => { const o = Object.fromEntries(db.prepare('SELECT key,value FROM settings').all().map(r => [r.key, r.value])); delete o.jwt_secret; delete o.invoice_seq; return o; };
const SECRET = process.env.JWT_SECRET || getS('jwt_secret') || (() => { const s = crypto.randomBytes(48).toString('hex'); setS('jwt_secret', s); return s; })();

// ───────── helpers ─────────
class HttpErr extends Error { constructor(s, m) { super(m); this.status = s; } }
const bad = (m, s = 400) => { throw new HttpErr(s, m); };
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const str = (v, max = 200) => String(v ?? '').trim().slice(0, max);
const num = (v, min = -Infinity, max = Infinity, label = 'Value') => { const n = Number(v); if (v === '' || v === null || v === undefined || !Number.isFinite(n) || n < min || n > max) bad(`${label} is invalid`); return n; };
const wrap = fn => (req, res, next) => { try { const r = fn(req, res, next); if (r && r.catch) r.catch(next); } catch (e) { next(e); } };
const offMin = () => parseInt(getS('tz_offset') ?? '330') || 0;
const offStr = () => (offMin() >= 0 ? '+' : '') + offMin() + ' minutes';
const shift = (d = new Date()) => new Date(d.getTime() + offMin() * 60000);
const today = () => shift().toISOString().slice(0, 10);
const addDays = (s, n) => { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const isDate = s => /^\d{4}-\d{2}-\d{2}$/.test(s || '');
const L = "date(datetime(o.created_at,@off))";
const A = (req, action, entity, id, detail) => db.prepare('INSERT INTO activity_logs(user_id,username,action,entity,entity_id,detail,ip) VALUES(?,?,?,?,?,?,?)')
  .run(req.user?.id ?? null, req.user?.username ?? null, action, entity ?? null, id == null ? null : String(id), detail ? String(detail).slice(0, 500) : null, req.ip);

// ───────── roles & permissions ─────────
const ALL = ['pos.sell', 'pos.discount', 'orders.view_all', 'orders.view_own', 'orders.cancel', 'products.view', 'products.manage', 'inventory.manage', 'customers.manage', 'creditors.manage', 'reports.view', 'reports.export', 'users.manage', 'logs.view'];
const DEF = {
  owner: ALL,
  manager: ['pos.sell', 'pos.discount', 'orders.view_all', 'orders.view_own', 'products.view', 'inventory.manage', 'customers.manage', 'creditors.manage', 'reports.view', 'reports.export', 'users.manage', 'logs.view'],
  employee: ['pos.sell', 'orders.view_own']
};
function effectivePerms(u) {
  if (u.role === 'owner') return ALL;
  let p = DEF[u.role];
  if (u.permissions) { try { const x = JSON.parse(u.permissions); if (Array.isArray(x)) p = x.filter(k => ALL.includes(k)); } catch {} }
  if (u.role !== 'manager') p = p.filter(k => k !== 'users.manage');
  return p;
}
const pubUser = u => ({ id: u.id, username: u.username, name: u.name, role: u.role, active: !!u.active, created_at: u.created_at, last_login: u.last_login, permissions: effectivePerms(u) });

// ───────── app ─────────
const app = express(); app.set('trust proxy', 1); app.disable('x-powered-by');
app.use((req, res, next) => { res.set({ 'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'SAMEORIGIN', 'Referrer-Policy': 'same-origin' }); next(); });
app.use(express.json({ limit: '2mb' })); app.use(cookieParser());
app.use('/api', (req, res, next) => { if (req.method !== 'GET' && req.get('X-CSRF') !== '1') return res.status(403).json({ error: 'Blocked (CSRF)' }); next(); });

function auth(req, res, next) {
  const t = req.cookies.token; if (!t) return res.status(401).json({ error: 'Please sign in' });
  try {
    const p = jwt.verify(t, SECRET); const u = db.prepare('SELECT * FROM users WHERE id=?').get(p.id);
    if (!u || !u.active) return res.status(401).json({ error: 'Account disabled' });
    req.user = u; req.perms = effectivePerms(u); next();
  } catch { res.status(401).json({ error: 'Session expired' }); }
}
const need = (...ps) => (req, res, next) => ps.some(p => req.perms.includes(p)) ? next() : res.status(403).json({ error: 'You do not have permission for this action' });
const ownerOnly = (req, res, next) => req.user.role === 'owner' ? next() : res.status(403).json({ error: 'Owner only' });

// SSE
const clients = new Set();
app.get('/api/events', auth, (req, res) => {
  res.set({ 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive', 'X-Accel-Buffering': 'no' }); res.flushHeaders();
  res.write('retry: 3000\n\n'); const c = { res, id: req.user.id, perms: req.perms }; clients.add(c); req.on('close', () => clients.delete(c));
});
function emit(type, data, ownerId) { for (const c of clients) { if (type === 'order' && !(c.perms.includes('orders.view_all') || c.id === ownerId)) continue; try { c.res.write(`event: ${type}\ndata: ${JSON.stringify(data)}\n\n`); } catch {} } }
setInterval(() => { for (const c of clients) try { c.res.write(': ping\n\n'); } catch {} }, 25000).unref();

// ───────── auth routes ─────────
const fails = new Map();
app.get('/api/status', (req, res) => res.json({ needsSetup: db.prepare('SELECT COUNT(*) n FROM users').get().n === 0, business: getS('business_name'), logo: getS('logo') }));
app.post('/api/setup', wrap((req, res) => {
  if (db.prepare('SELECT COUNT(*) n FROM users').get().n) bad('Setup already completed', 403);
  const b = req.body || {}, name = str(b.name, 80), username = str(b.username, 32), pw = String(b.password || '');
  if (!name || !/^[a-z0-9._-]{3,32}$/i.test(username)) bad('Enter a name and a username (3-32 letters/numbers)');
  if (pw.length < 8) bad('Password must be at least 8 characters');
  if (b.business) setS('business_name', str(b.business, 100));
  const id = db.prepare('INSERT INTO users(username,name,password_hash,role) VALUES(?,?,?,?)').run(username, name, bcrypt.hashSync(pw, 10), 'owner').lastInsertRowid;
  req.user = { id, username }; A(req, 'setup', 'user', id, 'Owner account created'); res.json({ ok: true });
}));
app.post('/api/login', wrap((req, res) => {
  const b = req.body || {}, username = str(b.username, 40), key = req.ip + '|' + username.toLowerCase(), f = fails.get(key);
  if (f && f.n >= 8 && Date.now() - f.t < 15 * 60000) bad('Too many attempts. Try again in 15 minutes.', 429);
  const u = db.prepare('SELECT * FROM users WHERE username=?').get(username);
  if (!u || !u.active || !bcrypt.compareSync(String(b.password || ''), u.password_hash)) {
    fails.set(key, { n: (f && Date.now() - f.t < 15 * 60000 ? f.n : 0) + 1, t: Date.now() }); req.user = u || { username }; A(req, 'login.failed', 'user', u?.id, username); bad('Invalid username or password', 401);
  }
  fails.delete(key); db.prepare("UPDATE users SET last_login=datetime('now') WHERE id=?").run(u.id); req.user = u; A(req, 'login', 'user', u.id);
  res.cookie('token', jwt.sign({ id: u.id }, SECRET, { expiresIn: '12h' }), { httpOnly: true, sameSite: 'lax', secure: req.secure, maxAge: 12 * 3600e3 });
  res.json({ ok: true });
}));
app.post('/api/logout', auth, (req, res) => { A(req, 'logout', 'user', req.user.id); res.clearCookie('token'); res.json({ ok: true }); });
app.get('/api/me', auth, (req, res) => res.json({ user: pubUser(req.user), perms: req.perms, settings: pubSettings() }));
app.patch('/api/me/password', auth, wrap((req, res) => {
  const b = req.body || {}; if (!bcrypt.compareSync(String(b.current || ''), req.user.password_hash)) bad('Current password is wrong');
  if (String(b.password || '').length < 8) bad('New password must be at least 8 characters');
  db.prepare('UPDATE users SET password_hash=? WHERE id=?').run(bcrypt.hashSync(b.password, 10), req.user.id); A(req, 'password.change', 'user', req.user.id); res.json({ ok: true });
}));
app.get('/api/me/stats', auth, (req, res) => res.json(db.prepare(`SELECT COALESCE(SUM(total),0) revenue, COUNT(*) orders FROM orders o WHERE user_id=@u AND status='completed' AND ${L}=@t`).get({ u: req.user.id, off: offStr(), t: today() })));

// ───────── users ─────────
function cleanPerms(req, role, list) {
  if (!Array.isArray(list)) return null;
  let p = list.filter(k => ALL.includes(k)); if (req.user.role !== 'owner') p = p.filter(k => req.perms.includes(k)); if (role !== 'manager') p = p.filter(k => k !== 'users.manage');
  return JSON.stringify([...new Set(p)]);
}
app.get('/api/users', auth, need('users.manage'), (req, res) => {
  const rows = req.user.role === 'owner' ? db.prepare('SELECT * FROM users ORDER BY id').all() : db.prepare("SELECT * FROM users WHERE role='employee' ORDER BY id").all();
  res.json({ users: rows.map(pubUser), defaults: DEF, all: ALL });
});
app.post('/api/users', auth, need('users.manage'), wrap((req, res) => {
  const b = req.body || {}, role = b.role, name = str(b.name, 80), username = str(b.username, 32);
  if (!['manager', 'employee'].includes(role)) bad('Choose a role'); if (req.user.role !== 'owner' && role !== 'employee') bad('Managers can only create employees', 403);
  if (!name || !/^[a-z0-9._-]{3,32}$/i.test(username)) bad('Name and username (3-32 letters/numbers/._-) are required');
  if (String(b.password || '').length < 8) bad('Password must be at least 8 characters');
  if (db.prepare('SELECT 1 FROM users WHERE username=?').get(username)) bad('Username already taken', 409);
  const id = db.prepare('INSERT INTO users(username,name,password_hash,role,permissions,created_by) VALUES(?,?,?,?,?,?)').run(username, name, bcrypt.hashSync(b.password, 10), role, cleanPerms(req, role, b.permissions), req.user.id).lastInsertRowid;
  A(req, 'user.create', 'user', id, `${username} (${role})`); res.json(pubUser(db.prepare('SELECT * FROM users WHERE id=?').get(id)));
}));
app.patch('/api/users/:id', auth, need('users.manage'), wrap((req, res) => {
  const t = db.prepare('SELECT * FROM users WHERE id=?').get(+req.params.id); if (!t) bad('User not found', 404);
  if (t.role === 'owner') bad('The owner account cannot be modified here', 403);
  if (req.user.role !== 'owner' && t.role !== 'employee') bad('Managers can only manage employees', 403);
  const b = req.body || {}; let role = t.role;
  if (b.role !== undefined && b.role !== t.role) { if (req.user.role !== 'owner') bad('Only the owner can change roles', 403); if (!['manager', 'employee'].includes(b.role)) bad('Invalid role'); role = b.role; }
  const name = b.name !== undefined ? str(b.name, 80) : t.name; if (!name) bad('Name required');
  let active = t.active; if (b.active !== undefined) { if (t.id === req.user.id) bad('You cannot deactivate yourself'); active = b.active ? 1 : 0; }
  const perms = b.permissions !== undefined ? cleanPerms(req, role, b.permissions) : (role !== t.role ? null : t.permissions);
  let hash = t.password_hash; if (b.password) { if (String(b.password).length < 8) bad('Password must be at least 8 characters'); hash = bcrypt.hashSync(b.password, 10); }
  db.prepare('UPDATE users SET name=?,role=?,active=?,permissions=?,password_hash=? WHERE id=?').run(name, role, active, perms, hash, t.id);
  if (!active) for (const c of clients) if (c.id === t.id) { try { c.res.end(); } catch {} clients.delete(c); }
  A(req, active !== t.active ? (active ? 'user.activate' : 'user.deactivate') : 'user.update', 'user', t.id, `${t.username}${b.password ? ' (password reset)' : ''}${role !== t.role ? ` role ${t.role}→${role}` : ''}`);
  res.json(pubUser(db.prepare('SELECT * FROM users WHERE id=?').get(t.id)));
}));

// ───────── products & stock ─────────
app.get('/api/products', auth, need('pos.sell', 'products.view'), (req, res) => {
  const full = req.perms.includes('products.view'), cost = req.perms.includes('products.manage') || req.perms.includes('reports.view');
  const rows = db.prepare(`SELECT id,sku,name,price,${cost ? 'cost,' : ''}tax_pct,discount_pct,stock,active FROM products ${full && req.query.all === '1' ? '' : 'WHERE active=1'} ORDER BY name`).all();
  res.json(rows);
});
function prodBody(b, cur) {
  const name = str(b.name ?? cur?.name, 120).toUpperCase(); if (!name) bad('Product name is required');
  return { name, price: num(b.price ?? cur?.price, 0, 1e9, 'Price'), cost: num(b.cost ?? cur?.cost ?? 0, 0, 1e9, 'Cost'), tax_pct: num(b.tax_pct ?? cur?.tax_pct ?? 0, 0, 100, 'Tax'), discount_pct: num(b.discount_pct ?? cur?.discount_pct ?? 0, 0, 100, 'Discount') };
}
app.post('/api/products', auth, need('products.manage'), wrap((req, res) => {
  const b = req.body || {}, p = prodBody(b), sku = str(b.sku, 40).toUpperCase() || 'SKU-' + Date.now().toString(36).toUpperCase();
  if (db.prepare('SELECT 1 FROM products WHERE sku=?').get(sku)) bad('SKU already exists', 409);
  const stock = num(b.stock ?? 0, 0, 1e9, 'Stock');
  const id = db.transaction(() => { const id = db.prepare('INSERT INTO products(sku,name,price,cost,tax_pct,discount_pct,stock) VALUES(?,?,?,?,?,?,?)').run(sku, p.name, p.price, p.cost, p.tax_pct, p.discount_pct, stock).lastInsertRowid;
    if (stock) db.prepare('INSERT INTO stock_movements(product_id,delta,reason,user_id) VALUES(?,?,?,?)').run(id, stock, 'Opening stock', req.user.id); return id; })();
  A(req, 'product.create', 'product', id, `${p.name} @ ${p.price}`); emit('product', { id }); res.json({ id });
}));
app.patch('/api/products/:id', auth, need('products.manage'), wrap((req, res) => {
  const cur = db.prepare('SELECT * FROM products WHERE id=?').get(+req.params.id); if (!cur) bad('Product not found', 404);
  const b = req.body || {}, p = prodBody(b, cur), active = b.active === undefined ? cur.active : (b.active ? 1 : 0);
  db.prepare("UPDATE products SET name=?,price=?,cost=?,tax_pct=?,discount_pct=?,active=?,updated_at=datetime('now') WHERE id=?").run(p.name, p.price, p.cost, p.tax_pct, p.discount_pct, active, cur.id);
  if (p.price !== cur.price) A(req, 'product.price', 'product', cur.id, `${cur.name}: ${cur.price} → ${p.price}`);
  if (p.discount_pct !== cur.discount_pct) A(req, 'product.discount', 'product', cur.id, `${cur.name}: ${cur.discount_pct}% → ${p.discount_pct}%`);
  if (active !== cur.active) A(req, active ? 'product.restore' : 'product.archive', 'product', cur.id, cur.name);
  if (p.name !== cur.name || p.cost !== cur.cost || p.tax_pct !== cur.tax_pct) A(req, 'product.update', 'product', cur.id, p.name);
  emit('product', { id: cur.id }); res.json({ ok: true });
}));
app.delete('/api/products/:id', auth, need('products.manage'), wrap((req, res) => {
  const cur = db.prepare('SELECT * FROM products WHERE id=?').get(+req.params.id); if (!cur) bad('Product not found', 404);
  db.prepare('UPDATE products SET active=0 WHERE id=?').run(cur.id); A(req, 'product.archive', 'product', cur.id, cur.name); emit('product', { id: cur.id }); res.json({ ok: true });
}));
app.post('/api/products/:id/stock', auth, need('inventory.manage'), wrap((req, res) => {
  const cur = db.prepare('SELECT * FROM products WHERE id=?').get(+req.params.id); if (!cur) bad('Product not found', 404);
  const delta = num(req.body?.delta, -1e9, 1e9, 'Quantity'); if (!delta) bad('Quantity cannot be zero'); if (cur.stock + delta < -1e-9) bad('Stock cannot go below zero');
  db.transaction(() => { db.prepare("UPDATE products SET stock=stock+?,updated_at=datetime('now') WHERE id=?").run(delta, cur.id); db.prepare('INSERT INTO stock_movements(product_id,delta,reason,user_id) VALUES(?,?,?,?)').run(cur.id, delta, str(req.body?.reason, 120) || 'Stock adjustment', req.user.id); })();
  A(req, 'stock.adjust', 'product', cur.id, `${cur.name}: ${delta > 0 ? '+' : ''}${delta}`); emit('product', { id: cur.id }); res.json({ stock: r2(cur.stock + delta) });
}));
app.get('/api/stock/movements', auth, need('inventory.manage'), (req, res) => res.json(db.prepare('SELECT m.*,p.name product,u.name user FROM stock_movements m JOIN products p ON p.id=m.product_id LEFT JOIN users u ON u.id=m.user_id ORDER BY m.id DESC LIMIT 100').all()));

// ───────── customers, creditors, payments ─────────
const BAL = `(COALESCE((SELECT SUM(total-paid) FROM orders WHERE customer_id=c.id AND status='completed'),0)-COALESCE((SELECT SUM(amount) FROM payments WHERE customer_id=c.id),0))`;
const custBalance = id => r2(db.prepare(`SELECT ${BAL} b FROM customers c WHERE c.id=?`).get(id)?.b || 0);
app.get('/api/customers', auth, need('pos.sell', 'customers.manage', 'creditors.manage'), (req, res) => {
  const q = `%${str(req.query.q, 60)}%`; res.json(db.prepare(`SELECT c.*,ROUND(${BAL},2) balance FROM customers c WHERE c.name LIKE ? OR c.phone LIKE ? ORDER BY c.name LIMIT 500`).all(q, q));
});
app.post('/api/customers', auth, need('pos.sell', 'customers.manage'), wrap((req, res) => {
  const name = str(req.body?.name, 100); if (!name) bad('Customer name is required');
  const id = db.prepare('INSERT INTO customers(name,phone,address) VALUES(?,?,?)').run(name, str(req.body?.phone, 30), str(req.body?.address, 200)).lastInsertRowid; A(req, 'customer.create', 'customer', id, name); res.json({ id, name });
}));
app.patch('/api/customers/:id', auth, need('customers.manage'), wrap((req, res) => {
  const c = db.prepare('SELECT * FROM customers WHERE id=?').get(+req.params.id); if (!c) bad('Customer not found', 404);
  const name = str(req.body?.name ?? c.name, 100); if (!name) bad('Name required');
  db.prepare('UPDATE customers SET name=?,phone=?,address=? WHERE id=?').run(name, str(req.body?.phone ?? c.phone, 30), str(req.body?.address ?? c.address, 200), c.id); A(req, 'customer.update', 'customer', c.id, name); res.json({ ok: true });
}));
app.get('/api/customers/:id/ledger', auth, need('customers.manage', 'creditors.manage'), wrap((req, res) => {
  const c = db.prepare('SELECT * FROM customers WHERE id=?').get(+req.params.id); if (!c) bad('Customer not found', 404);
  res.json({ customer: c, balance: custBalance(c.id), orders: db.prepare("SELECT id,invoice_no,created_at,total,paid,status,payment_method FROM orders WHERE customer_id=? ORDER BY id DESC LIMIT 100").all(c.id),
    payments: db.prepare('SELECT p.*,u.name user FROM payments p JOIN users u ON u.id=p.user_id WHERE customer_id=? ORDER BY p.id DESC LIMIT 100').all(c.id) });
}));
app.get('/api/creditors', auth, need('creditors.manage', 'reports.view'), (req, res) => res.json(db.prepare(`SELECT * FROM (SELECT c.id,c.name,c.phone,ROUND(${BAL},2) balance,(SELECT MAX(created_at) FROM orders WHERE customer_id=c.id) last_order FROM customers c) WHERE balance>0.005 ORDER BY balance DESC`).all()));
app.post('/api/payments', auth, need('creditors.manage'), wrap((req, res) => {
  const c = db.prepare('SELECT * FROM customers WHERE id=?').get(+req.body?.customer_id); if (!c) bad('Customer not found', 404);
  const amt = r2(num(req.body?.amount, 0.01, 1e9, 'Amount')), bal = custBalance(c.id); if (amt > bal + 0.005) bad(`Amount exceeds outstanding balance (${bal.toFixed(2)})`);
  const method = ['cash', 'card', 'transfer'].includes(req.body?.method) ? req.body.method : 'cash';
  const id = db.prepare('INSERT INTO payments(customer_id,amount,method,note,user_id) VALUES(?,?,?,?,?)').run(c.id, amt, method, str(req.body?.note, 200), req.user.id).lastInsertRowid;
  A(req, 'payment.create', 'customer', c.id, `${c.name}: ${amt} (${method})`); emit('payment', { customer_id: c.id }); res.json({ id, balance: custBalance(c.id) });
}));

// ───────── orders ─────────
function orderFull(id) {
  const o = db.prepare('SELECT o.*,u.name employee,c.phone customer_phone,c.address customer_address FROM orders o JOIN users u ON u.id=o.user_id LEFT JOIN customers c ON c.id=o.customer_id WHERE o.id=?').get(id);
  if (o) o.items = db.prepare('SELECT * FROM order_items WHERE order_id=? ORDER BY id').all(id); return o;
}
function canSee(req, o) { return req.perms.includes('orders.view_all') || (req.perms.includes('orders.view_own') && o.user_id === req.user.id); }
const gen = new Map();
function ensurePdf(id) {
  const o = orderFull(id), file = path.join(INV, `${o.invoice_no}.pdf`); if (fs.existsSync(file)) return Promise.resolve(file);
  if (gen.has(file)) return gen.get(file);
  const p = new Promise((ok, no) => { const tmp = file + '.tmp', ws = fs.createWriteStream(tmp); const d = invoicePDF(o, pubSettings(), 'a4'); d.pipe(ws);
    ws.on('finish', () => { fs.renameSync(tmp, file); ok(file); }); ws.on('error', no); }).finally(() => gen.delete(file));
  gen.set(file, p); return p;
}
app.post('/api/orders', auth, need('pos.sell'), wrap((req, res) => {
  const b = req.body || {}, ref = str(b.client_ref, 80); if (!/^[\w-]{8,80}$/.test(ref)) bad('Missing order reference');
  if (!Array.isArray(b.items) || !b.items.length || b.items.length > 200) bad('Add at least one item');
  const method = ['cash', 'card', 'transfer', 'credit'].includes(b.payment_method) ? b.payment_method : bad('Choose a payment method');
  const canDisc = req.perms.includes('pos.discount');
  const out = db.transaction(() => {
    const ex = db.prepare('SELECT id,invoice_no,total FROM orders WHERE client_ref=?').get(ref); if (ex) return { ...ex, duplicate: true };
    let cust = null; if (b.customer_id) { cust = db.prepare('SELECT * FROM customers WHERE id=?').get(+b.customer_id); if (!cust) bad('Customer not found'); }
    if (method === 'credit' && !cust) bad('Select a customer for credit sales');
    const lines = []; let sub = 0, disc = 0, tax = 0, total = 0;
    for (const it of b.items) {
      const p = db.prepare('SELECT * FROM products WHERE id=? AND active=1').get(+it.product_id); if (!p) bad('A product is no longer available');
      const qty = num(it.qty, 0.001, 100000, 'Quantity'), ret = it.is_return ? 1 : 0, sg = ret ? -1 : 1;
      let dp = p.discount_pct; if (canDisc && it.discount_pct !== undefined && it.discount_pct !== null && it.discount_pct !== '') dp = num(it.discount_pct, 0, 100, 'Discount');
      const gross = r2(p.price * qty), da = r2(gross * dp / 100), ta = r2((gross - da) * p.tax_pct / 100), lt = r2(gross - da + ta);
      const up = ret ? db.prepare("UPDATE products SET stock=stock+?,updated_at=datetime('now') WHERE id=?").run(qty, p.id)
        : db.prepare("UPDATE products SET stock=stock-@q,updated_at=datetime('now') WHERE id=@id AND stock>=@q-0.0000001").run({ q: qty, id: p.id });
      if (!up.changes) bad(`Not enough stock for ${p.name} (available ${db.prepare('SELECT stock FROM products WHERE id=?').get(p.id).stock})`, 409);
      sub += sg * gross; disc += sg * da; tax += sg * ta; total += sg * lt;
      lines.push({ p, qty, ret, dp, da, ta, lt: sg * lt });
    }
    sub = r2(sub); disc = r2(disc); tax = r2(tax); total = r2(total); if (total < 0) bad('Returns exceed the sale total');
    const paid = method === 'credit' ? r2(num(b.paid ?? 0, 0, total, 'Paid amount')) : total;
    const prev = cust ? Math.max(0, custBalance(cust.id)) : 0;
    const seq = (parseInt(getS('invoice_seq') || '0') || 0) + 1; setS('invoice_seq', seq); const inv = 'INV-' + String(seq).padStart(6, '0');
    const id = db.prepare('INSERT INTO orders(invoice_no,client_ref,customer_id,customer_name,user_id,subtotal,discount_total,tax_total,total,paid,prev_outstanding,payment_method,note) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)')
      .run(inv, ref, cust?.id ?? null, cust?.name ?? (str(b.customer_name, 100) || 'Walk-in Customer'), req.user.id, sub, disc, tax, total, paid, prev, method, str(b.note, 200)).lastInsertRowid;
    const ins = db.prepare('INSERT INTO order_items(order_id,product_id,sku,name,qty,unit_price,cost,discount_pct,discount_amt,tax_pct,tax_amt,line_total,is_return) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)');
    lines.forEach(l => ins.run(id, l.p.id, l.p.sku, l.p.name, l.qty, l.p.price, l.p.cost, l.dp, l.da, l.p.tax_pct, l.ta, l.lt, l.ret));
    return { id, invoice_no: inv, total, disc };
  })();
  if (out.duplicate) return res.json(out);
  A(req, 'order.create', 'order', out.id, `${out.invoice_no} Rs.${out.total} ${method}`); if (out.disc > 0) A(req, 'discount.applied', 'order', out.id, `${out.invoice_no}: Rs.${out.disc}`);
  ensurePdf(out.id).catch(e => logErr(e, 'pdf'));
  emit('order', { id: out.id, invoice_no: out.invoice_no, total: out.total, employee: req.user.name, event: 'created' }, req.user.id); emit('product', {});
  res.json({ id: out.id, invoice_no: out.invoice_no, total: out.total });
}));
app.get('/api/orders', auth, need('orders.view_all', 'orders.view_own'), (req, res) => {
  const w = ['1=1'], p = { off: offStr() }, all = req.perms.includes('orders.view_all');
  if (!all) { w.push('o.user_id=@me'); p.me = req.user.id; } else if (+req.query.user) { w.push('o.user_id=@u'); p.u = +req.query.user; }
  if (isDate(req.query.from)) { w.push(`${L}>=@from`); p.from = req.query.from; } if (isDate(req.query.to)) { w.push(`${L}<=@to`); p.to = req.query.to; }
  if (['completed', 'cancelled'].includes(req.query.status)) { w.push('o.status=@st'); p.st = req.query.status; }
  if (req.query.q) { w.push('(o.invoice_no LIKE @q OR o.customer_name LIKE @q)'); p.q = `%${str(req.query.q, 40)}%`; }
  res.json(db.prepare(`SELECT o.id,o.invoice_no,o.created_at,o.customer_name,o.total,o.paid,o.discount_total,o.payment_method,o.status,u.name employee FROM orders o JOIN users u ON u.id=o.user_id WHERE ${w.join(' AND ')} ORDER BY o.id DESC LIMIT ${Math.min(+req.query.limit || 200, 500)}`).all(p));
});
app.get('/api/orders/:id', auth, need('orders.view_all', 'orders.view_own'), wrap((req, res) => { const o = orderFull(+req.params.id); if (!o || !canSee(req, o)) bad('Order not found', 404); res.json(o); }));
app.get('/api/orders/:id/pdf', auth, need('orders.view_all', 'orders.view_own'), wrap(async (req, res) => {
  const o = orderFull(+req.params.id); if (!o || !canSee(req, o)) bad('Order not found', 404);
  res.set({ 'Content-Type': 'application/pdf', 'Content-Disposition': `${req.query.download ? 'attachment' : 'inline'}; filename="${o.invoice_no}${req.query.format === 'thermal' ? '-57mm' : ''}.pdf"` });
  if (req.query.format === 'thermal') return invoicePDF(o, pubSettings(), 'thermal').pipe(res);
  res.sendFile(await ensurePdf(o.id));
}));
app.post('/api/orders/:id/cancel', auth, need('orders.cancel'), wrap((req, res) => {
  const reason = str(req.body?.reason, 200); if (!reason) bad('A cancellation reason is required');
  const o = db.prepare('SELECT * FROM orders WHERE id=?').get(+req.params.id); if (!o) bad('Order not found', 404); if (o.status === 'cancelled') bad('Order already cancelled', 409);
  db.transaction(() => {
    for (const i of db.prepare('SELECT * FROM order_items WHERE order_id=?').all(o.id)) { const d = i.is_return ? -i.qty : i.qty; db.prepare('UPDATE products SET stock=stock+? WHERE id=?').run(d, i.product_id); }
    db.prepare("UPDATE orders SET status='cancelled',cancelled_at=datetime('now'),cancelled_by=?,cancel_reason=? WHERE id=?").run(req.user.id, reason, o.id);
  })();
  try { fs.unlinkSync(path.join(INV, `${o.invoice_no}.pdf`)); } catch {}
  ensurePdf(o.id).catch(e => logErr(e, 'pdf')); A(req, 'order.cancel', 'order', o.id, `${o.invoice_no}: ${reason}`);
  emit('order', { id: o.id, invoice_no: o.invoice_no, event: 'cancelled' }, o.user_id); emit('product', {}); res.json({ ok: true });
}));

// ───────── analytics ─────────
function summary(from, to) {
  const off = offStr(), t = today(), P = (a, b) => db.prepare(`SELECT COALESCE(SUM(total),0) revenue,COUNT(*) orders,COALESCE(SUM(discount_total),0) discounts FROM orders o WHERE status='completed' AND ${L} BETWEEN @a AND @b`).get({ off, a, b });
  const cards = { today: P(t, t), yesterday: P(addDays(t, -1), addDays(t, -1)), week: P(addDays(t, -6), t), month: P(t.slice(0, 8) + '01', t), year: P(t.slice(0, 5) + '01-01', t) };
  const base = { off, a: from, b: to };
  const range_stats = db.prepare(`SELECT COALESCE(SUM(total),0) revenue,COUNT(*) orders,COALESCE(SUM(discount_total),0) discounts,COALESCE(SUM(tax_total),0) tax,COALESCE(SUM(CASE WHEN paid>=total-0.005 THEN 1 ELSE 0 END),0) paid_orders,COALESCE(SUM(CASE WHEN paid<total-0.005 THEN 1 ELSE 0 END),0) unpaid_orders FROM orders o WHERE status='completed' AND ${L} BETWEEN @a AND @b`).get(base);
  range_stats.avg_order = range_stats.orders ? r2(range_stats.revenue / range_stats.orders) : 0; range_stats.cancelled = db.prepare(`SELECT COUNT(*) n FROM orders o WHERE status='cancelled' AND ${L} BETWEEN @a AND @b`).get(base).n;
  const grp = e => db.prepare(`SELECT ${e} label,COALESCE(SUM(total),0) revenue,COUNT(*) orders FROM orders o WHERE status='completed' AND ${L} BETWEEN @a AND @b GROUP BY label ORDER BY label`).all(base);
  const day = grp(L), map = new Map(day.map(r => [r.label, r])), filled = []; for (let d = from, n = 0; d <= to && n < 400; d = addDays(d, 1), n++) filled.push(map.get(d) || { label: d, revenue: 0, orders: 0 });
  const series = { day: filled, week: grp(`strftime('%Y-W%W',${L})`), month: grp(`substr(${L},1,7)`), year: grp(`substr(${L},1,4)`) };
  const products = db.prepare(`SELECT p.id,p.name,COALESCE(a.qty,0) qty,COALESCE(a.revenue,0) revenue,COALESCE(a.discount,0) discount FROM products p LEFT JOIN (SELECT oi.product_id pid,SUM(CASE WHEN oi.is_return THEN -oi.qty ELSE oi.qty END) qty,SUM(oi.line_total) revenue,SUM(CASE WHEN oi.is_return THEN -oi.discount_amt ELSE oi.discount_amt END) discount FROM order_items oi JOIN orders o ON o.id=oi.order_id WHERE o.status='completed' AND ${L} BETWEEN @a AND @b GROUP BY oi.product_id) a ON a.pid=p.id WHERE p.active=1 ORDER BY qty DESC,revenue DESC`).all(base);
  const employees = db.prepare(`SELECT u.id,u.name,u.role,COUNT(o.id) orders,COALESCE(SUM(o.total),0) revenue,COALESCE(SUM(o.discount_total),0) discounts FROM users u LEFT JOIN orders o ON o.user_id=u.id AND o.status='completed' AND ${L} BETWEEN @a AND @b GROUP BY u.id ORDER BY revenue DESC`).all(base).filter(e => e.role !== 'owner' || e.orders);
  const payments = db.prepare(`SELECT payment_method method,COUNT(*) orders,COALESCE(SUM(total),0) revenue FROM orders o WHERE status='completed' AND ${L} BETWEEN @a AND @b GROUP BY payment_method ORDER BY revenue DESC`).all(base);
  const outstanding = r2(db.prepare(`SELECT COALESCE(SUM(b),0) v FROM (SELECT ${BAL} b FROM customers c) WHERE b>0.005`).get().v);
  return { today: t, range: { from, to }, cards, range_stats, series, products, employees, payments, outstanding };
}
function rangeOf(q) { const t = today(); let to = isDate(q.to) ? q.to : t, from = isDate(q.from) ? q.from : addDays(to, -29); if (from > to) [from, to] = [to, from]; if (new Date(to) - new Date(from) > 3660 * 864e5) bad('Range too large'); return [from, to]; }
app.get('/api/reports/summary', auth, need('reports.view'), (req, res) => res.json(summary(...rangeOf(req.query))));
app.get('/api/reports/export', auth, need('reports.export'), wrap(async (req, res) => {
  const [from, to] = rangeOf(req.query), d = summary(from, to), S = pubSettings(); A(req, 'report.export', 'report', null, `${req.query.type || 'xlsx'} ${from}..${to}`);
  if (req.query.type === 'pdf') { res.set({ 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="report-${from}_${to}.pdf"` }); return reportPDF(d, S).pipe(res); }
  const wb = new ExcelJS.Workbook(); wb.creator = S.business_name; const hd = ws => { ws.getRow(1).font = { bold: true }; ws.views = [{ state: 'frozen', ySplit: 1 }]; };
  let ws = wb.addWorksheet('Summary'); ws.columns = [{ header: 'Metric', width: 34 }, { header: 'Value', width: 20 }]; hd(ws); const r = d.range_stats;
  [['From', from], ['To', to], ['Total revenue', r.revenue], ['Orders', r.orders], ['Average order value', r.avg_order], ['Total discounts', r.discounts], ['Tax collected', r.tax], ['Paid orders', r.paid_orders], ['Unpaid/partial orders', r.unpaid_orders], ['Cancelled orders', r.cancelled], ['Outstanding creditor balance', d.outstanding]].forEach(x => ws.addRow(x));
  ws = wb.addWorksheet('Daily'); ws.columns = [{ header: 'Date', width: 14 }, { header: 'Orders', width: 10 }, { header: 'Revenue', width: 16 }]; hd(ws); d.series.day.forEach(s => ws.addRow([s.label, s.orders, s.revenue]));
  ws = wb.addWorksheet('Products'); ws.columns = [{ header: 'Product', width: 34 }, { header: 'Qty sold', width: 12 }, { header: 'Revenue', width: 16 }, { header: 'Discount', width: 14 }]; hd(ws); d.products.forEach(p => ws.addRow([p.name, p.qty, p.revenue, p.discount]));
  ws = wb.addWorksheet('Employees'); ws.columns = [{ header: 'Name', width: 26 }, { header: 'Role', width: 12 }, { header: 'Orders', width: 10 }, { header: 'Revenue', width: 16 }, { header: 'Discounts', width: 14 }]; hd(ws); d.employees.forEach(e => ws.addRow([e.name, e.role, e.orders, e.revenue, e.discounts]));
  ws = wb.addWorksheet('Orders'); ws.columns = ['Invoice', 'Date', 'Employee', 'Customer', 'Method', 'Subtotal', 'Discount', 'Tax', 'Total', 'Paid', 'Status'].map(h => ({ header: h, width: 16 })); hd(ws);
  db.prepare(`SELECT o.*,u.name employee,datetime(o.created_at,@off) local FROM orders o JOIN users u ON u.id=o.user_id WHERE ${L} BETWEEN @a AND @b ORDER BY o.id`).all({ off: offStr(), a: from, b: to }).forEach(o => ws.addRow([o.invoice_no, o.local, o.employee, o.customer_name, o.payment_method, o.subtotal, o.discount_total, o.tax_total, o.total, o.paid, o.status]));
  res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="report-${from}_${to}.xlsx"` }); await wb.xlsx.write(res); res.end();
}));

// ───────── logs, settings, backups ─────────
app.get('/api/logs', auth, need('logs.view'), (req, res) => {
  const q = `%${str(req.query.q, 40)}%`, own = req.user.role === 'owner';
  res.json(db.prepare(`SELECT l.*,u.role FROM activity_logs l LEFT JOIN users u ON u.id=l.user_id WHERE (l.action LIKE @q OR l.username LIKE @q OR l.detail LIKE @q) ${own ? '' : "AND COALESCE(u.role,'')!='owner'"} ORDER BY l.id DESC LIMIT 300`).all({ q }));
});
app.patch('/api/settings', auth, ownerOnly, wrap((req, res) => {
  const b = req.body || {}, keys = ['business_name', 'business_address', 'business_phone', 'receipt_footer', 'route_name', 'tz_offset', 'low_stock', 'logo'];
  for (const k of keys) if (b[k] !== undefined) {
    let v = String(b[k]);
    if (k === 'logo') { if (v && (!/^data:image\/(png|jpeg);base64,/.test(v) || v.length > 700000)) bad('Logo must be a PNG/JPEG under 500 KB'); }
    else if (k === 'tz_offset') v = String(num(v, -720, 840, 'Time zone offset')); else if (k === 'low_stock') v = String(num(v, 0, 1e6, 'Low stock')); else v = str(v, 200);
    setS(k, v);
  }
  try { for (const f of fs.readdirSync(INV)) fs.unlinkSync(path.join(INV, f)); } catch {} // invoices re-render with new branding on demand
  A(req, 'settings.update', 'settings', null, Object.keys(b).filter(k => keys.includes(k)).join(',')); res.json(pubSettings());
}));
const bkName = () => 'billease-' + new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14) + '.db';
async function makeBackup(prefix = '') { const n = prefix + bkName(); await db.backup(path.join(BK, n)); return n; }
const listBk = () => fs.readdirSync(BK).filter(f => f.endsWith('.db')).map(f => ({ name: f, size: fs.statSync(path.join(BK, f)).size, created: fs.statSync(path.join(BK, f)).mtime })).sort((a, b) => b.created - a.created);
app.get('/api/backups', auth, ownerOnly, (req, res) => res.json(listBk()));
app.post('/api/backups', auth, ownerOnly, wrap(async (req, res) => { const n = await makeBackup(); A(req, 'backup.create', 'backup', n); res.json({ name: n }); }));
app.get('/api/backups/:name', auth, ownerOnly, wrap((req, res) => { const n = req.params.name; if (!/^[\w.-]+\.db$/.test(n) || !fs.existsSync(path.join(BK, n))) bad('Backup not found', 404); A(req, 'backup.download', 'backup', n); res.download(path.join(BK, n)); }));
app.post('/api/backups/:name/restore', auth, ownerOnly, wrap((req, res) => {
  const n = req.params.name; if (!/^[\w.-]+\.db$/.test(n) || !fs.existsSync(path.join(BK, n))) bad('Backup not found', 404);
  A(req, 'backup.restore', 'backup', n); fs.copyFileSync(path.join(BK, n), PEND); res.json({ ok: true, restarting: true }); setTimeout(() => process.exit(0), 700);
}));
async function autoBackup() { try { const pre = 'auto-' + today().replace(/-/g, ''); if (!listBk().some(b => b.name.startsWith(pre))) { await db.backup(path.join(BK, pre + '-' + bkName())); }
  const autos = listBk().filter(b => b.name.startsWith('auto-')); autos.slice(14).forEach(b => fs.unlinkSync(path.join(BK, b.name))); } catch (e) { logErr(e, 'autobackup'); } }
if (!process.env.NO_AUTOBACKUP) { setTimeout(autoBackup, 15000).unref(); setInterval(autoBackup, 3600e3).unref(); }

// ───────── static & errors ─────────
app.get('/vendor/chart.js', (req, res) => res.sendFile(path.join(__dirname, 'node_modules/chart.js/dist/chart.umd.js')));
app.use(express.static(path.join(__dirname, 'public'), { maxAge: 0 }));
app.use('/api', (req, res) => res.status(404).json({ error: 'Not found' }));
app.use((e, req, res, next) => {
  if (e instanceof HttpErr) return res.status(e.status).json({ error: e.message });
  if (e && /SQLITE_CONSTRAINT/.test(e.code || '')) { logErr(e, req.originalUrl); return res.status(409).json({ error: 'Conflict: duplicate or invalid data' }); }
  logErr(e, req.method + ' ' + req.originalUrl); if (!res.headersSent) res.status(500).json({ error: 'Server error. It has been logged.' });
});
const PORT = process.env.PORT || 3000;
if (require.main === module) app.listen(PORT, () => console.log(`BillEase running on http://localhost:${PORT}`));
module.exports = app;
