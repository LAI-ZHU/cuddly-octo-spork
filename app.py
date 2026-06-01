"""服装进销存管理系统"""
import sqlite3, os, sys, io
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'data.db')
app = Flask(__name__)
app.secret_key = 'fuzhuang_2026'

# ─── DB ───
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            style_no TEXT NOT NULL UNIQUE,
            name TEXT DEFAULT '',
            color TEXT DEFAULT '',
            season TEXT DEFAULT '',
            supplier TEXT DEFAULT '',
            cost_price REAL DEFAULT 0,
            retail_price REAL DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            size TEXT NOT NULL,
            quantity INTEGER DEFAULT 0,
            batch TEXT DEFAULT '',
            updated_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(product_id) REFERENCES products(id),
            UNIQUE(product_id, size, batch)
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            size TEXT NOT NULL,
            type TEXT NOT NULL CHECK(type IN ('进货','出货','退货')),
            quantity INTEGER NOT NULL,
            customer TEXT DEFAULT '',
            batch TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(product_id) REFERENCES products(id)
        );
    ''')
    # 兼容旧表：加价格列
    try: conn.execute('ALTER TABLE products ADD COLUMN cost_price REAL DEFAULT 0')
    except: pass
    try: conn.execute('ALTER TABLE products ADD COLUMN retail_price REAL DEFAULT 0')
    except: pass
    conn.commit()
    conn.close()

init_db()

def get_sizes():
    return ['120', '130', '140', '150', '160']

def get_product_summary():
    conn = get_db()
    rows = conn.execute('''
        SELECT p.*, COALESCE(SUM(i.quantity),0) as total_stock
        FROM products p
        LEFT JOIN inventory i ON i.product_id = p.id
        GROUP BY p.id ORDER BY p.style_no
    ''').fetchall()
    conn.close()
    return rows

# ─── 路由 ───
@app.route('/')
def index():
    products = get_product_summary()
    # 全局统计
    conn = get_db()
    total_cost = 0
    total_retail = 0
    for p in products:
        if p['cost_price']:
            total_cost += p['cost_price'] * p['total_stock']
        if p['retail_price']:
            total_retail += p['retail_price'] * p['total_stock']
    conn.close()
    return render_template('index.html', products=products, sizes=get_sizes(),
                          total_cost=total_cost, total_retail=total_retail)

@app.route('/product/add', methods=['GET','POST'])
def add_product():
    if request.method == 'POST':
        style_no = request.form['style_no'].strip()
        name = request.form.get('name','').strip()
        color = request.form.get('color','').strip()
        season = request.form.get('season','').strip()
        supplier = request.form.get('supplier','').strip()
        cost_price = float(request.form.get('cost_price',0) or 0)
        retail_price = float(request.form.get('retail_price',0) or 0)
        if not style_no:
            flash('款号不能为空', 'error')
            return redirect(url_for('add_product'))
        conn = get_db()
        try:
            conn.execute('INSERT INTO products (style_no,name,color,season,supplier,cost_price,retail_price) VALUES (?,?,?,?,?,?,?)',
                        (style_no,name,color,season,supplier,cost_price,retail_price))
            conn.commit()
            flash(f'款号 {style_no} 已添加', 'success')
        except sqlite3.IntegrityError:
            flash(f'款号 {style_no} 已存在', 'error')
        finally:
            conn.close()
        return redirect(url_for('index'))
    return render_template('add_product.html')

@app.route('/product/<int:pid>', methods=['GET','POST'])
def product_detail(pid):
    conn = get_db()
    p = conn.execute('SELECT * FROM products WHERE id=?', (pid,)).fetchone()
    if not p:
        flash('款号不存在', 'error')
        conn.close()
        return redirect(url_for('index'))
    # 更新价格
    if request.method == 'POST':
        cost = float(request.form.get('cost_price',0) or 0)
        retail = float(request.form.get('retail_price',0) or 0)
        conn.execute('UPDATE products SET cost_price=?, retail_price=? WHERE id=?', (cost, retail, pid))
        conn.commit()
        flash('价格已更新', 'success')
        conn.close()
        return redirect(url_for('product_detail', pid=pid))
    inv = conn.execute('SELECT * FROM inventory WHERE product_id=? ORDER BY batch DESC, size', (pid,)).fetchall()
    txs = conn.execute('SELECT * FROM transactions WHERE product_id=? ORDER BY created_at DESC LIMIT 200', (pid,)).fetchall()
    conn.close()
    return render_template('detail.html', product=p, inv=inv, txs=txs, sizes=get_sizes())

@app.route('/stock/in', methods=['GET','POST'])
def stock_in():
    conn = get_db()
    products = conn.execute('SELECT id,style_no,name,color,cost_price,retail_price FROM products ORDER BY style_no').fetchall()
    conn.close()
    if request.method == 'POST':
        pid = request.form['product_id']
        size = request.form['size']
        qty = int(request.form['quantity'])
        batch = request.form.get('batch','').strip()
        notes = request.form.get('notes','').strip()
        if qty <= 0:
            flash('数量必须大于0', 'error')
            return redirect(url_for('stock_in'))
        conn = get_db()
        conn.execute('''
            INSERT INTO inventory (product_id,size,quantity,batch)
            VALUES (?,?,?,?)
            ON CONFLICT(product_id,size,batch) DO UPDATE SET quantity=quantity+?, updated_at=datetime('now','localtime')
        ''', (pid,size,qty,batch,qty))
        conn.execute('INSERT INTO transactions (product_id,size,type,quantity,batch,notes) VALUES (?,?,?,?,?,?)',
                    (pid,size,'进货',qty,batch,notes))
        conn.commit()
        conn.close()
        flash('进货已录入', 'success')
        return redirect(url_for('index'))
    return render_template('stock_in.html', products=products, sizes=get_sizes())

@app.route('/stock/out', methods=['GET','POST'])
def stock_out():
    conn = get_db()
    products = conn.execute('SELECT id,style_no,name,color,cost_price,retail_price FROM products ORDER BY style_no').fetchall()
    conn.close()
    if request.method == 'POST':
        pid = request.form['product_id']
        size = request.form['size']
        qty = int(request.form['quantity'])
        customer = request.form.get('customer','').strip()
        batch = request.form.get('batch','').strip()
        notes = request.form.get('notes','').strip()
        if qty <= 0:
            flash('数量必须大于0', 'error')
            return redirect(url_for('stock_out'))
        conn = get_db()
        inv_rows = conn.execute('''
            SELECT * FROM inventory WHERE product_id=? AND size=? AND quantity>0 ORDER BY updated_at DESC
        ''', (pid,size)).fetchall()
        remaining = qty
        for row in inv_rows:
            if remaining <= 0: break
            deduct = min(remaining, row['quantity'])
            conn.execute('UPDATE inventory SET quantity=quantity-? WHERE id=?', (deduct, row['id']))
            remaining -= deduct
        shipped = qty - remaining
        if remaining > 0:
            flash(f'库存不足，仅出货 {shipped} 件，缺 {remaining} 件', 'warning')
        conn.execute('INSERT INTO transactions (product_id,size,type,quantity,customer,batch,notes) VALUES (?,?,?,?,?,?,?)',
                    (pid,size,'出货',shipped,customer,batch,notes))
        conn.commit()
        conn.close()
        flash(f'出货已录入 {shipped} 件', 'success')
        return redirect(url_for('index'))
    return render_template('stock_out.html', products=products, sizes=get_sizes())

@app.route('/stock/return', methods=['GET','POST'])
def stock_return():
    conn = get_db()
    products = conn.execute('SELECT id,style_no,name,color,cost_price,retail_price FROM products ORDER BY style_no').fetchall()
    conn.close()
    if request.method == 'POST':
        pid = request.form['product_id']
        size = request.form['size']
        qty = int(request.form['quantity'])
        batch = request.form.get('batch','').strip()
        notes = request.form.get('notes','').strip()
        if qty <= 0:
            flash('数量必须大于0', 'error')
            return redirect(url_for('stock_return'))
        conn = get_db()
        conn.execute('''
            INSERT INTO inventory (product_id,size,quantity,batch)
            VALUES (?,?,?,?)
            ON CONFLICT(product_id,size,batch) DO UPDATE SET quantity=quantity+?, updated_at=datetime('now','localtime')
        ''', (pid,size,qty,batch,qty))
        conn.execute('INSERT INTO transactions (product_id,size,type,quantity,batch,notes) VALUES (?,?,?,?,?,?)',
                    (pid,size,'退货',qty,batch,notes))
        conn.commit()
        conn.close()
        flash('退货已录入', 'success')
        return redirect(url_for('index'))
    return render_template('stock_return.html', products=products, sizes=get_sizes())

@app.route('/report')
def report():
    conn = get_db()
    tx_type = request.args.get('type','')
    date_from = request.args.get('from','')
    date_to = request.args.get('to','')
    sql = '''
        SELECT t.*, p.style_no, p.name as product_name, p.color, p.cost_price, p.retail_price
        FROM transactions t
        JOIN products p ON p.id = t.product_id
        WHERE 1=1
    '''
    params = []
    if tx_type:
        sql += ' AND t.type=?'
        params.append(tx_type)
    if date_from:
        sql += ' AND t.created_at>=?'
        params.append(date_from)
    if date_to:
        sql += ' AND t.created_at<=?'
        params.append(date_to + ' 23:59:59')
    sql += ' ORDER BY t.created_at DESC LIMIT 500'
    txs = conn.execute(sql, params).fetchall()
    
    # 今日统计（含金额）
    stats = conn.execute('''
        SELECT type, SUM(quantity) as total_qty,
               COALESCE(SUM(t.quantity * p.retail_price), 0) as total_amount
        FROM transactions t
        JOIN products p ON p.id = t.product_id
        WHERE DATE(t.created_at) = DATE('now','localtime')
        GROUP BY t.type
    ''').fetchall()
    conn.close()
    return render_template('report.html', txs=txs, stats=stats,
                           tx_type=tx_type, date_from=date_from, date_to=date_to)

@app.route('/api/low-stock')
def low_stock():
    threshold = request.args.get('t', 5, type=int)
    conn = get_db()
    rows = conn.execute('''
        SELECT p.style_no, p.name, p.color, i.size, i.quantity
        FROM inventory i JOIN products p ON p.id = i.product_id
        WHERE i.quantity > 0 AND i.quantity <= ?
        ORDER BY i.quantity
    ''', (threshold,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
