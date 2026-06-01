"""服装进销存管理系统 v2 — 完整版（对标秦丝/货宝宝）"""
import sqlite3, os, sys, io, json, re, random
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import xlsxwriter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'data.db')
app = Flask(__name__)
app.secret_key = 'fuzhuang_v2_2026'
SIZES = ['均码','100','110','120','130','140','150','160','165','170','175','180']

# ═══════════════════════ DB ═══════════════════════
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
            parent_id INTEGER DEFAULT NULL, sort_order INTEGER DEFAULT 0,
            FOREIGN KEY(parent_id) REFERENCES categories(id)
        );
        CREATE TABLE IF NOT EXISTS brands (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE, remark TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT, style_no TEXT NOT NULL UNIQUE,
            name TEXT DEFAULT '', category_id INTEGER, brand_id INTEGER,
            color TEXT DEFAULT '', unit TEXT DEFAULT '件', barcode TEXT DEFAULT '',
            cost_price REAL DEFAULT 0, retail_price REAL DEFAULT 0, wholesale_price REAL DEFAULT 0,
            image TEXT DEFAULT '', description TEXT DEFAULT '', status TEXT DEFAULT '上架',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS product_sizes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL,
            size TEXT NOT NULL,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE,
            UNIQUE(product_id, size)
        );
        CREATE TABLE IF NOT EXISTS warehouses (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE,
            address TEXT DEFAULT '', remark TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL,
            size TEXT NOT NULL, warehouse_id INTEGER DEFAULT 1, quantity INTEGER DEFAULT 0,
            FOREIGN KEY(product_id) REFERENCES products(id),
            UNIQUE(product_id, size, warehouse_id)
        );
        CREATE TABLE IF NOT EXISTS inventory_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL,
            size TEXT NOT NULL, warehouse_id INTEGER DEFAULT 1,
            batch_no TEXT DEFAULT '', quantity INTEGER DEFAULT 0, cost_price REAL DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
            contact TEXT DEFAULT '', phone TEXT DEFAULT '', address TEXT DEFAULT '',
            remark TEXT DEFAULT '', balance REAL DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS purchase_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT, order_no TEXT NOT NULL UNIQUE,
            supplier_id INTEGER, total_amount REAL DEFAULT 0, paid_amount REAL DEFAULT 0,
            status TEXT DEFAULT '已采购', notes TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS purchase_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL, size TEXT DEFAULT '',
            quantity INTEGER DEFAULT 0, unit_price REAL DEFAULT 0, subtotal REAL DEFAULT 0,
            FOREIGN KEY(order_id) REFERENCES purchase_orders(id),
            FOREIGN KEY(product_id) REFERENCES products(id)
        );
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
            phone TEXT DEFAULT '', wechat TEXT DEFAULT '', address TEXT DEFAULT '',
            credit_limit REAL DEFAULT 0, balance REAL DEFAULT 0,
            level TEXT DEFAULT '普通会员', remark TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS sales_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT, order_no TEXT NOT NULL UNIQUE,
            customer_id INTEGER, total_amount REAL DEFAULT 0, discount REAL DEFAULT 0,
            final_amount REAL DEFAULT 0, payment_method TEXT DEFAULT '现金',
            status TEXT DEFAULT '已完成', notes TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS sales_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL, size TEXT DEFAULT '',
            quantity INTEGER DEFAULT 0, unit_price REAL DEFAULT 0,
            cost_price REAL DEFAULT 0, subtotal REAL DEFAULT 0,
            FOREIGN KEY(order_id) REFERENCES sales_orders(id),
            FOREIGN KEY(product_id) REFERENCES products(id)
        );
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT DEFAULT '',
            amount REAL DEFAULT 0, description TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS stock_counts (
            id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL,
            size TEXT DEFAULT '', warehouse_id INTEGER DEFAULT 1,
            expected_qty INTEGER DEFAULT 0, actual_qty INTEGER DEFAULT 0,
            difference INTEGER DEFAULT 0, notes TEXT DEFAULT '',
            count_date TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(product_id) REFERENCES products(id)
        );
    ''')
    # 默认数据
    if not conn.execute('SELECT id FROM warehouses LIMIT 1').fetchone():
        conn.execute("INSERT INTO warehouses (name,remark) VALUES ('主仓库','默认仓库')")
    conn.commit()
    conn.close()

init_db()

# ═══════════════════════ 辅助函数 ═══════════════════════
def gen_order_no(prefix):
    today = date.today().strftime('%Y%m%d')
    conn = get_db()
    row = conn.execute("SELECT COUNT(*) as c FROM sales_orders WHERE order_no LIKE ?",
                       (f'{prefix}{today}%',)).fetchone()
    conn.close()
    n = row['c'] + 1
    return f'{prefix}{today}-{n:03d}'

def get_product_stock(pid, size='', wid=1):
    conn = get_db()
    if size:
        row = conn.execute('SELECT SUM(quantity) as q FROM inventory WHERE product_id=? AND size=? AND warehouse_id=?',
                          (pid,size,wid)).fetchone()
    else:
        row = conn.execute('SELECT SUM(quantity) as q FROM inventory WHERE product_id=? AND warehouse_id=?',
                          (pid,wid)).fetchone()
    conn.close()
    return row['q'] or 0

# ═══════════════════════ 路由：首页 / 仪表盘 ═══════════════════════
@app.route('/')
def dashboard():
    conn = get_db()
    today_s = date.today().strftime('%Y-%m-%d')
    # 今日销售
    row = conn.execute("""
        SELECT COUNT(*) as orders, COALESCE(SUM(final_amount),0) as amount,
               COALESCE(SUM(final_amount - si.cost_price * si.quantity),0) as profit
        FROM sales_orders so
        JOIN sales_items si ON si.order_id = so.id
        WHERE DATE(so.created_at)=? AND so.status='已完成'
    """, (today_s,)).fetchone()
    today_sales = dict(row) if row['orders'] else {'orders':0,'amount':0,'profit':0}
    # 今日采购
    row2 = conn.execute("""
        SELECT COUNT(*) as orders, COALESCE(SUM(total_amount),0) as amount
        FROM purchase_orders WHERE DATE(created_at)=?
    """, (today_s,)).fetchone()
    today_purchase = dict(row2) if row2['orders'] else {'orders':0,'amount':0}
    # 总商品数/库存
    total_products = conn.execute('SELECT COUNT(*) as c FROM products').fetchone()['c']
    total_stock = conn.execute('SELECT COALESCE(SUM(quantity),0) as c FROM inventory').fetchone()['c']
    # 供应商/客户
    supplier_count = conn.execute('SELECT COUNT(*) as c FROM suppliers').fetchone()['c']
    customer_count = conn.execute('SELECT COUNT(*) as c FROM customers').fetchone()['c']
    # 低库存
    low_stock = conn.execute("""
        SELECT p.style_no, p.name, p.color, i.size, i.quantity
        FROM inventory i JOIN products p ON p.id=i.product_id
        WHERE i.quantity>0 AND i.quantity<=5 ORDER BY i.quantity LIMIT 20
    """).fetchall()
    # 欠款客户
    debt_customers = conn.execute("""
        SELECT id, name, balance FROM customers WHERE balance>0 ORDER BY balance DESC LIMIT 10
    """).fetchall()
    # 近7天销售额
    days = []
    for i in range(6, -1, -1):
        d = (date.today() - timedelta(days=i)).strftime('%Y-%m-%d')
        r = conn.execute("""
            SELECT COALESCE(SUM(final_amount),0) as amt FROM sales_orders
            WHERE DATE(created_at)=? AND status='已完成'
        """, (d,)).fetchone()
        days.append({'date': d[-5:], 'amount': r['amt']})
    # 本月销售额
    month_start = date.today().replace(day=1).strftime('%Y-%m-%d')
    month_sales = conn.execute("""
        SELECT COALESCE(SUM(final_amount),0) as amt FROM sales_orders
        WHERE created_at>=? AND status='已完成'
    """, (month_start,)).fetchone()['amt']
    conn.close()
    return render_template('dashboard.html', today_sales=today_sales, today_purchase=today_purchase,
                          total_products=total_products, total_stock=total_stock,
                          supplier_count=supplier_count, customer_count=customer_count,
                          low_stock=low_stock, debt_customers=debt_customers,
                          chart_days=days, month_sales=month_sales)

# ═══════════════════════ 路由：商品 ═══════════════════════
@app.route('/products')
def product_list():
    conn = get_db()
    cat_id = request.args.get('cat','')
    kw = request.args.get('kw','')
    sql = '''
        SELECT p.*, c.name as category_name, b.name as brand_name,
               (SELECT COALESCE(SUM(quantity),0) FROM inventory WHERE product_id=p.id) as stock
        FROM products p
        LEFT JOIN categories c ON c.id=p.category_id
        LEFT JOIN brands b ON b.id=p.brand_id
        WHERE 1=1
    '''
    params = []
    if cat_id:
        sql += ' AND p.category_id=?'
        params.append(cat_id)
    if kw:
        sql += ' AND (p.style_no LIKE ? OR p.name LIKE ? OR p.barcode LIKE ?)'
        kwp = f'%{kw}%'
        params.extend([kwp, kwp, kwp])
    sql += ' ORDER BY p.created_at DESC'
    products = conn.execute(sql, params).fetchall()
    categories = conn.execute('SELECT * FROM categories ORDER BY parent_id, sort_order').fetchall()
    conn.close()
    return render_template('product_list.html', products=products, categories=categories, cat_id=cat_id, kw=kw)

@app.route('/products/add', methods=['GET','POST'])
def product_add():
    conn = get_db()
    categories = conn.execute('SELECT * FROM categories ORDER BY parent_id, sort_order').fetchall()
    brands = conn.execute('SELECT * FROM brands ORDER BY name').fetchall()
    conn.close()
    if request.method == 'POST':
        style_no = request.form['style_no'].strip()
        if not style_no:
            flash('款号不能为空', 'error')
            return redirect(url_for('product_add'))
        name = request.form.get('name','').strip()
        cat_id = request.form.get('category_id') or None
        brand_id = request.form.get('brand_id') or None
        color = request.form.get('color','').strip()
        unit = request.form.get('unit','件')
        barcode = request.form.get('barcode','').strip()
        cost_price = float(request.form.get('cost_price',0) or 0)
        retail_price = float(request.form.get('retail_price',0) or 0)
        wholesale_price = float(request.form.get('wholesale_price',0) or 0)
        description = request.form.get('description','').strip()
        sizes = request.form.getlist('sizes[]')
        conn = get_db()
        try:
            conn.execute('''INSERT INTO products (style_no,name,category_id,brand_id,color,unit,barcode,
                cost_price,retail_price,wholesale_price,description) VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                (style_no,name,cat_id,brand_id,color,unit,barcode,cost_price,retail_price,wholesale_price,description))
            pid = conn.execute('SELECT id FROM products WHERE style_no=?', (style_no,)).fetchone()['id']
            for s in sizes:
                if s:
                    conn.execute('INSERT OR IGNORE INTO product_sizes (product_id,size) VALUES (?,?)', (pid,s))
            conn.commit()
            flash(f'商品 {style_no} 已添加', 'success')
        except sqlite3.IntegrityError:
            flash(f'款号 {style_no} 已存在', 'error')
        finally:
            conn.close()
        return redirect(url_for('product_list'))
    return render_template('product_form.html', product=None, categories=categories, brands=brands, sizes=SIZES)

@app.route('/products/<int:pid>', methods=['GET','POST'])
def product_edit(pid):
    conn = get_db()
    p = conn.execute('SELECT * FROM products WHERE id=?', (pid,)).fetchone()
    if not p:
        flash('商品不存在', 'error')
        conn.close()
        return redirect(url_for('product_list'))
    if request.method == 'POST':
        style_no = request.form['style_no'].strip()
        name = request.form.get('name','').strip()
        cat_id = request.form.get('category_id') or None
        brand_id = request.form.get('brand_id') or None
        color = request.form.get('color','').strip()
        unit = request.form.get('unit','件')
        barcode = request.form.get('barcode','').strip()
        cost_price = float(request.form.get('cost_price',0) or 0)
        retail_price = float(request.form.get('retail_price',0) or 0)
        wholesale_price = float(request.form.get('wholesale_price',0) or 0)
        description = request.form.get('description','').strip()
        status = request.form.get('status','上架')
        sizes = request.form.getlist('sizes[]')
        try:
            conn.execute('''UPDATE products SET style_no=?,name=?,category_id=?,brand_id=?,color=?,unit=?,
                barcode=?,cost_price=?,retail_price=?,wholesale_price=?,description=?,status=?,
                updated_at=datetime('now','localtime') WHERE id=?''',
                (style_no,name,cat_id,brand_id,color,unit,barcode,cost_price,retail_price,wholesale_price,description,status,pid))
            conn.execute('DELETE FROM product_sizes WHERE product_id=?', (pid,))
            for s in sizes:
                if s:
                    conn.execute('INSERT OR IGNORE INTO product_sizes (product_id,size) VALUES (?,?)', (pid,s))
            conn.commit()
            flash('商品已更新', 'success')
        except sqlite3.IntegrityError:
            flash(f'款号 {style_no} 已被占用', 'error')
        conn.close()
        return redirect(url_for('product_list'))
    categories = conn.execute('SELECT * FROM categories ORDER BY parent_id, sort_order').fetchall()
    brands = conn.execute('SELECT * FROM brands ORDER BY name').fetchall()
    psizes = [r['size'] for r in conn.execute('SELECT size FROM product_sizes WHERE product_id=?', (pid,)).fetchall()]
    stocks = conn.execute('SELECT size, quantity FROM inventory WHERE product_id=?', (pid,)).fetchall()
    stock_map = {r['size']: r['quantity'] for r in stocks}
    conn.close()
    return render_template('product_form.html', product=p, categories=categories, brands=brands,
                          sizes=SIZES, psizes=psizes, stock_map=stock_map)

@app.route('/products/<int:pid>/delete', methods=['POST'])
def product_delete(pid):
    conn = get_db()
    conn.execute('DELETE FROM products WHERE id=?', (pid,))
    conn.execute('DELETE FROM inventory WHERE product_id=?', (pid,))
    conn.commit()
    conn.close()
    flash('商品已删除', 'success')
    return redirect(url_for('product_list'))

@app.route('/products/<int:pid>/stock')
def product_stock(pid):
    conn = get_db()
    p = conn.execute('SELECT * FROM products WHERE id=?', (pid,)).fetchone()
    if not p:
        conn.close()
        return jsonify([])
    batches = conn.execute('''
        SELECT ib.*, w.name as warehouse_name
        FROM inventory_batches ib
        JOIN warehouses w ON w.id=ib.warehouse_id
        WHERE ib.product_id=? AND ib.quantity>0
        ORDER BY ib.created_at DESC
    ''', (pid,)).fetchall()
    stocks = conn.execute('SELECT size, warehouse_id, quantity FROM inventory WHERE product_id=?', (pid,)).fetchall()
    conn.close()
    return jsonify({'batches': [dict(r) for r in batches], 'stocks': [dict(r) for r in stocks]})

# ═══════════════════════ 分类 ═══════════════════════
@app.route('/categories')
def category_list():
    conn = get_db()
    cats = conn.execute('''
        SELECT c.*, (SELECT COUNT(*) FROM products WHERE category_id=c.id) as product_count
        FROM categories c ORDER BY c.parent_id, c.sort_order
    ''').fetchall()
    conn.close()
    return render_template('category_list.html', categories=cats)

@app.route('/categories/add', methods=['POST'])
def category_add():
    name = request.form['name'].strip()
    parent_id = request.form.get('parent_id') or None
    if not name:
        flash('分类名不能为空', 'error')
        return redirect(url_for('category_list'))
    conn = get_db()
    conn.execute('INSERT INTO categories (name,parent_id) VALUES (?,?)', (name,parent_id))
    conn.commit()
    conn.close()
    flash(f'分类 {name} 已添加', 'success')
    return redirect(url_for('category_list'))

@app.route('/categories/<int:cid>/delete', methods=['POST'])
def category_delete(cid):
    conn = get_db()
    conn.execute('DELETE FROM categories WHERE id=?', (cid,))
    conn.commit()
    conn.close()
    flash('分类已删除', 'success')
    return redirect(url_for('category_list'))

# ═══════════════════════ 品牌 ═══════════════════════
@app.route('/brands')
def brand_list():
    conn = get_db()
    brands = conn.execute('''
        SELECT b.*, (SELECT COUNT(*) FROM products WHERE brand_id=b.id) as product_count
        FROM brands b ORDER BY b.name
    ''').fetchall()
    conn.close()
    return render_template('brand_list.html', brands=brands)

@app.route('/brands/add', methods=['POST'])
def brand_add():
    name = request.form['name'].strip()
    if not name:
        flash('品牌名不能为空', 'error')
        return redirect(url_for('brand_list'))
    conn = get_db()
    try:
        conn.execute('INSERT INTO brands (name) VALUES (?)', (name,))
        conn.commit()
        flash(f'品牌 {name} 已添加', 'success')
    except sqlite3.IntegrityError:
        flash('品牌已存在', 'error')
    conn.close()
    return redirect(url_for('brand_list'))

@app.route('/brands/<int:bid>/delete', methods=['POST'])
def brand_delete(bid):
    conn = get_db()
    conn.execute('DELETE FROM brands WHERE id=?', (bid,))
    conn.commit()
    conn.close()
    flash('品牌已删除', 'success')
    return redirect(url_for('brand_list'))

# ═══════════════════════ 采购 ═══════════════════════
@app.route('/purchase')
def purchase_list():
    conn = get_db()
    orders = conn.execute('''
        SELECT po.*, COALESCE(s.name,'-') as supplier_name
        FROM purchase_orders po
        LEFT JOIN suppliers s ON s.id=po.supplier_id
        ORDER BY po.created_at DESC LIMIT 200
    ''').fetchall()
    conn.close()
    return render_template('purchase_list.html', orders=orders)

@app.route('/purchase/add', methods=['GET','POST'])
def purchase_add():
    conn = get_db()
    suppliers = conn.execute('SELECT * FROM suppliers ORDER BY name').fetchall()
    products = conn.execute('SELECT id,style_no,name,color,cost_price,retail_price FROM products WHERE status="上架" ORDER BY style_no').fetchall()
    conn.close()
    if request.method == 'POST':
        supplier_id = request.form.get('supplier_id') or None
        notes = request.form.get('notes','').strip()
        pids = request.form.getlist('product_id[]')
        sizes = request.form.getlist('size[]')
        qtys = request.form.getlist('qty[]')
        prices = request.form.getlist('price[]')
        total = 0.0
        conn = get_db()
        order_no = gen_order_no('CG')
        conn.execute('INSERT INTO purchase_orders (order_no,supplier_id,notes) VALUES (?,?,?)',
                    (order_no,supplier_id,notes))
        oid = conn.execute('SELECT id FROM purchase_orders WHERE order_no=?', (order_no,)).fetchone()['id']
        for i in range(len(pids)):
            if not pids[i] or not qtys[i]: continue
            qty = int(qtys[i])
            price = float(prices[i] or 0)
            size = sizes[i] if i < len(sizes) else ''
            subtotal = qty * price
            conn.execute('INSERT INTO purchase_items (order_id,product_id,size,quantity,unit_price,subtotal) VALUES (?,?,?,?,?,?)',
                        (oid, pids[i], size, qty, price, subtotal))
            # 增加库存
            wid = 1
            # 批次
            batch_no = datetime.now().strftime('%Y%m%d') + f'-B{random.randint(1,99):02d}'
            conn.execute('INSERT INTO inventory_batches (product_id,size,warehouse_id,batch_no,quantity,cost_price) VALUES (?,?,?,?,?,?)',
                        (pids[i], size, wid, batch_no, qty, price))
            conn.execute('''
                INSERT INTO inventory (product_id,size,warehouse_id,quantity)
                VALUES (?,?,?,?)
                ON CONFLICT(product_id,size,warehouse_id) DO UPDATE SET quantity=quantity+?
            ''', (pids[i], size, wid, qty, qty))
            total += subtotal
        conn.execute('UPDATE purchase_orders SET total_amount=? WHERE id=?', (total, oid))
        conn.commit()
        conn.close()
        flash(f'采购单 {order_no} 已创建，总金额 ¥{total:.2f}', 'success')
        return redirect(url_for('purchase_list'))
    return render_template('purchase_form.html', suppliers=suppliers, products=products)

@app.route('/purchase/<int:oid>')
def purchase_detail(oid):
    conn = get_db()
    order = conn.execute('''
        SELECT po.*, COALESCE(s.name,'-') as supplier_name
        FROM purchase_orders po LEFT JOIN suppliers s ON s.id=po.supplier_id
        WHERE po.id=?
    ''', (oid,)).fetchone()
    if not order:
        conn.close()
        flash('采购单不存在', 'error')
        return redirect(url_for('purchase_list'))
    items = conn.execute('''
        SELECT pi.*, p.style_no, p.name, p.color
        FROM purchase_items pi JOIN products p ON p.id=pi.product_id
        WHERE pi.order_id=?
    ''', (oid,)).fetchall()
    conn.close()
    return render_template('purchase_detail.html', order=order, items=items)

@app.route('/purchase/<int:oid>/delete', methods=['POST'])
def purchase_delete(oid):
    conn = get_db()
    # 还原库存
    items = conn.execute('SELECT * FROM purchase_items WHERE order_id=?', (oid,)).fetchall()
    wid = 1
    for item in items:
        conn.execute('''
            UPDATE inventory SET quantity=MAX(quantity-?,0)
            WHERE product_id=? AND size=? AND warehouse_id=?
        ''', (item['quantity'], item['product_id'], item['size'], wid))
    conn.execute('DELETE FROM purchase_items WHERE order_id=?', (oid,))
    conn.execute('DELETE FROM purchase_orders WHERE id=?', (oid,))
    conn.commit()
    conn.close()
    flash('采购单已删除', 'success')
    return redirect(url_for('purchase_list'))

# ═══════════════════════ 销售 ═══════════════════════
@app.route('/sales')
def sales_list():
    conn = get_db()
    orders = conn.execute('''
        SELECT so.*, COALESCE(c.name,'-') as customer_name
        FROM sales_orders so LEFT JOIN customers c ON c.id=so.customer_id
        ORDER BY so.created_at DESC LIMIT 200
    ''').fetchall()
    conn.close()
    return render_template('sales_list.html', orders=orders)

@app.route('/sales/quick')
def sales_quick():
    """快速开单页(POS风格)"""
    conn = get_db()
    products = conn.execute('''
        SELECT p.*, COALESCE(SUM(i.quantity),0) as stock
        FROM products p LEFT JOIN inventory i ON i.product_id=p.id
        WHERE p.status='上架'
        GROUP BY p.id ORDER BY p.style_no
    ''').fetchall()
    customers = conn.execute('SELECT * FROM customers ORDER BY name').fetchall()
    # 每个商品的尺码库存
    skus = {}
    for p in products:
        szs = conn.execute('SELECT size, quantity FROM inventory WHERE product_id=?', (p['id'],)).fetchall()
        skus[p['id']] = {r['size']: r['quantity'] for r in szs}
    conn.close()
    return render_template('sales_quick.html', products=products, customers=customers, skus=skus)

@app.route('/sales/submit', methods=['POST'])
def sales_submit():
    data = request.get_json()
    if not data or not data.get('items'):
        return jsonify({'ok': False, 'msg': '没有商品'})
    conn = get_db()
    customer_id = data.get('customer_id') or None
    discount = float(data.get('discount', 0))
    payment = data.get('payment_method', '现金')
    notes = data.get('notes','').strip()
    order_no = gen_order_no('XS')
    total_amt = 0.0
    total_cost = 0.0
    conn.execute('INSERT INTO sales_orders (order_no,customer_id,discount,payment_method,notes) VALUES (?,?,?,?,?)',
                (order_no,customer_id,discount,payment,notes))
    oid = conn.execute('SELECT id FROM sales_orders WHERE order_no=?', (order_no,)).fetchone()['id']
    for item in data['items']:
        pid = item['product_id']
        size = item.get('size','')
        qty = int(item['quantity'])
        price = float(item['price'])
        # 取当时的进货价
        p = conn.execute('SELECT cost_price FROM products WHERE id=?', (pid,)).fetchone()
        cost = p['cost_price'] if p else 0
        subtotal = qty * price
        total_amt += subtotal
        total_cost += qty * cost
        conn.execute('INSERT INTO sales_items (order_id,product_id,size,quantity,unit_price,cost_price,subtotal) VALUES (?,?,?,?,?,?,?)',
                    (oid,pid,size,qty,price,cost,subtotal))
        # 扣减库存（从最新批次扣）
        remaining = qty
        batches = conn.execute('''
            SELECT * FROM inventory_batches WHERE product_id=? AND size=? AND quantity>0 ORDER BY created_at DESC
        ''', (pid, size)).fetchall()
        for b in batches:
            if remaining <= 0: break
            deduct = min(remaining, b['quantity'])
            conn.execute('UPDATE inventory_batches SET quantity=quantity-? WHERE id=?', (deduct, b['id']))
            remaining -= deduct
        conn.execute('''
            UPDATE inventory SET quantity=MAX(quantity-?,0)
            WHERE product_id=? AND size=? AND warehouse_id=1
        ''', (qty, pid, size))
    final_amt = max(0, total_amt - discount)
    conn.execute('UPDATE sales_orders SET total_amount=?,final_amount=? WHERE id=?',
                (total_amt, final_amt, oid))
    # 如果赊账，增加客户欠款
    if payment == '赊账' and customer_id:
        conn.execute('UPDATE customers SET balance=balance+? WHERE id=?', (final_amt, customer_id))
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'order_no': order_no, 'total': total_amt, 'final': final_amt})

@app.route('/sales/<int:oid>')
def sales_detail(oid):
    conn = get_db()
    order = conn.execute('''
        SELECT so.*, COALESCE(c.name,'-') as customer_name
        FROM sales_orders so LEFT JOIN customers c ON c.id=so.customer_id
        WHERE so.id=?
    ''', (oid,)).fetchone()
    if not order:
        conn.close()
        flash('销售单不存在', 'error')
        return redirect(url_for('sales_list'))
    items = conn.execute('''
        SELECT si.*, p.style_no, p.name, p.color
        FROM sales_items si JOIN products p ON p.id=si.product_id
        WHERE si.order_id=?
    ''', (oid,)).fetchall()
    conn.close()
    return render_template('sales_detail.html', order=order, items=items)

@app.route('/sales/<int:oid>/delete', methods=['POST'])
def sales_delete(oid):
    conn = get_db()
    items = conn.execute('SELECT * FROM sales_items WHERE order_id=?', (oid,)).fetchall()
    wid = 1
    for item in items:
        conn.execute('''
            UPDATE inventory SET quantity=quantity+?
            WHERE product_id=? AND size=? AND warehouse_id=?
        ''', (item['quantity'], item['product_id'], item['size'], wid))
    conn.execute('DELETE FROM sales_items WHERE order_id=?', (oid,))
    order = conn.execute('SELECT * FROM sales_orders WHERE id=?', (oid,)).fetchone()
    if order and order['payment_method'] == '赊账' and order['customer_id']:
        conn.execute('UPDATE customers SET balance=MAX(balance-?,0) WHERE id=?',
                    (order['final_amount'], order['customer_id']))
    conn.execute('DELETE FROM sales_orders WHERE id=?', (oid,))
    conn.commit()
    conn.close()
    flash('销售单已删除', 'success')
    return redirect(url_for('sales_list'))

# ═══════════════════════ 客户 ═══════════════════════
@app.route('/customers')
def customer_list():
    conn = get_db()
    kw = request.args.get('kw','')
    sql = 'SELECT *, (SELECT COUNT(*) FROM sales_orders WHERE customer_id=c.id) as order_count FROM customers c'
    params = []
    if kw:
        sql += ' WHERE c.name LIKE ? OR c.phone LIKE ?'
        k = f'%{kw}%'
        params.extend([k,k])
    sql += ' ORDER BY c.name'
    customers = conn.execute(sql, params).fetchall()
    conn.close()
    return render_template('customer_list.html', customers=customers, kw=kw)

@app.route('/customers/add', methods=['GET','POST'])
def customer_add():
    if request.method == 'POST':
        name = request.form['name'].strip()
        phone = request.form.get('phone','').strip()
        wechat = request.form.get('wechat','').strip()
        address = request.form.get('address','').strip()
        credit_limit = float(request.form.get('credit_limit',0) or 0)
        level = request.form.get('level','普通会员')
        remark = request.form.get('remark','').strip()
        if not name:
            flash('客户名不能为空', 'error')
            return redirect(url_for('customer_add'))
        conn = get_db()
        conn.execute('INSERT INTO customers (name,phone,wechat,address,credit_limit,level,remark) VALUES (?,?,?,?,?,?,?)',
                    (name,phone,wechat,address,credit_limit,level,remark))
        conn.commit()
        conn.close()
        flash(f'客户 {name} 已添加', 'success')
        return redirect(url_for('customer_list'))
    return render_template('customer_form.html', customer=None)

@app.route('/customers/<int:cid>', methods=['GET','POST'])
def customer_edit(cid):
    conn = get_db()
    c = conn.execute('SELECT * FROM customers WHERE id=?', (cid,)).fetchone()
    if not c:
        conn.close()
        flash('客户不存在', 'error')
        return redirect(url_for('customer_list'))
    if request.method == 'POST':
        name = request.form['name'].strip()
        phone = request.form.get('phone','').strip()
        wechat = request.form.get('wechat','').strip()
        address = request.form.get('address','').strip()
        credit_limit = float(request.form.get('credit_limit',0) or 0)
        level = request.form.get('level','普通会员')
        remark = request.form.get('remark','').strip()
        conn.execute('''UPDATE customers SET name=?,phone=?,wechat=?,address=?,credit_limit=?,level=?,remark=?
            WHERE id=?''', (name,phone,wechat,address,credit_limit,level,remark,cid))
        conn.commit()
        flash('客户已更新', 'success')
        conn.close()
        return redirect(url_for('customer_list'))
    orders = conn.execute('''
        SELECT * FROM sales_orders WHERE customer_id=? ORDER BY created_at DESC LIMIT 50
    ''', (cid,)).fetchall()
    conn.close()
    return render_template('customer_form.html', customer=c, orders=orders)

@app.route('/customers/<int:cid>/pay', methods=['POST'])
def customer_pay(cid):
    """客户还款"""
    amount = float(request.form.get('amount',0))
    if amount <= 0:
        flash('金额必须大于0', 'error')
        return redirect(url_for('customer_edit', cid=cid))
    conn = get_db()
    conn.execute('UPDATE customers SET balance=MAX(balance-?,0) WHERE id=?', (amount, cid))
    conn.commit()
    conn.close()
    flash(f'已收款 ¥{amount:.2f}', 'success')
    return redirect(url_for('customer_edit', cid=cid))

# ═══════════════════════ 供应商 ═══════════════════════
@app.route('/suppliers')
def supplier_list():
    conn = get_db()
    kw = request.args.get('kw','')
    sql = 'SELECT *, (SELECT COUNT(*) FROM purchase_orders WHERE supplier_id=s.id) as order_count FROM suppliers s'
    params = []
    if kw:
        sql += ' WHERE s.name LIKE ? OR s.contact LIKE ?'
        k = f'%{kw}%'
        params.extend([k,k])
    sql += ' ORDER BY s.name'
    suppliers = conn.execute(sql, params).fetchall()
    conn.close()
    return render_template('supplier_list.html', suppliers=suppliers, kw=kw)

@app.route('/suppliers/add', methods=['GET','POST'])
def supplier_add():
    if request.method == 'POST':
        name = request.form['name'].strip()
        contact = request.form.get('contact','').strip()
        phone = request.form.get('phone','').strip()
        address = request.form.get('address','').strip()
        remark = request.form.get('remark','').strip()
        if not name:
            flash('供应商名不能为空', 'error')
            return redirect(url_for('supplier_add'))
        conn = get_db()
        conn.execute('INSERT INTO suppliers (name,contact,phone,address,remark) VALUES (?,?,?,?,?)',
                    (name,contact,phone,address,remark))
        conn.commit()
        conn.close()
        flash(f'供应商 {name} 已添加', 'success')
        return redirect(url_for('supplier_list'))
    return render_template('supplier_form.html', supplier=None)

@app.route('/suppliers/<int:sid>', methods=['GET','POST'])
def supplier_edit(sid):
    conn = get_db()
    s = conn.execute('SELECT * FROM suppliers WHERE id=?', (sid,)).fetchone()
    if not s:
        conn.close()
        flash('供应商不存在', 'error')
        return redirect(url_for('supplier_list'))
    if request.method == 'POST':
        name = request.form['name'].strip()
        contact = request.form.get('contact','').strip()
        phone = request.form.get('phone','').strip()
        address = request.form.get('address','').strip()
        remark = request.form.get('remark','').strip()
        conn.execute('UPDATE suppliers SET name=?,contact=?,phone=?,address=?,remark=? WHERE id=?',
                    (name,contact,phone,address,remark,sid))
        conn.commit()
        flash('供应商已更新', 'success')
        conn.close()
        return redirect(url_for('supplier_list'))
    orders = conn.execute('SELECT * FROM purchase_orders WHERE supplier_id=? ORDER BY created_at DESC LIMIT 50', (sid,)).fetchall()
    conn.close()
    return render_template('supplier_form.html', supplier=s, orders=orders)

# ═══════════════════════ 库存 ═══════════════════════
@app.route('/inventory')
def inventory_overview():
    conn = get_db()
    wid = request.args.get('wid', 1)
    cat_id = request.args.get('cat','')
    kw = request.args.get('kw','')
    sql = '''
        SELECT p.id, p.style_no, p.name, p.color, p.cost_price, p.retail_price,
               i.size, i.quantity, c.name as category_name
        FROM inventory i
        JOIN products p ON p.id=i.product_id
        LEFT JOIN categories c ON c.id=p.category_id
        WHERE i.warehouse_id=?
    '''
    params = [wid]
    if cat_id:
        sql += ' AND p.category_id=?'
        params.append(cat_id)
    if kw:
        sql += ' AND (p.style_no LIKE ? OR p.name LIKE ?)'
        k = f'%{kw}%'
        params.extend([k,k])
    sql += ' ORDER BY p.style_no, i.size'
    inv = conn.execute(sql, params).fetchall()
    warehouses = conn.execute('SELECT * FROM warehouses').fetchall()
    categories = conn.execute('SELECT * FROM categories ORDER BY parent_id, sort_order').fetchall()
    conn.close()
    return render_template('inventory_overview.html', inv=inv, warehouses=warehouses, wid=int(wid),
                          categories=categories, cat_id=cat_id, kw=kw)

@app.route('/inventory/batches')
def inventory_batches():
    conn = get_db()
    pid = request.args.get('pid','')
    sql = '''
        SELECT ib.*, p.style_no, p.name, p.color, w.name as warehouse_name
        FROM inventory_batches ib
        JOIN products p ON p.id=ib.product_id
        JOIN warehouses w ON w.id=ib.warehouse_id
        WHERE ib.quantity>0
    '''
    params = []
    if pid:
        sql += ' AND ib.product_id=?'
        params.append(pid)
    sql += ' ORDER BY ib.created_at DESC LIMIT 500'
    batches = conn.execute(sql, params).fetchall()
    products = conn.execute('SELECT id, style_no, name FROM products ORDER BY style_no').fetchall()
    conn.close()
    return render_template('inventory_batches.html', batches=batches, products=products, pid=pid)

@app.route('/inventory/count', methods=['GET','POST'])
def inventory_count():
    conn = get_db()
    if request.method == 'POST':
        pid = request.form['product_id']
        size = request.form.get('size','')
        actual = int(request.form.get('actual_qty',0))
        expected = int(request.form.get('expected_qty',0))
        notes = request.form.get('notes','').strip()
        wid = request.form.get('warehouse_id', 1)
        diff = actual - expected
        conn.execute('''INSERT INTO stock_counts (product_id,size,warehouse_id,expected_qty,actual_qty,difference,notes)
            VALUES (?,?,?,?,?,?,?)''', (pid,size,wid,expected,actual,diff,notes))
        # 更新库存
        conn.execute('''
            UPDATE inventory SET quantity=?
            WHERE product_id=? AND size=? AND warehouse_id=?
        ''', (actual, pid, size, wid))
        conn.commit()
        flash(f'盘点完成，差异 {diff:+d}', 'success')
        conn.close()
        return redirect(url_for('inventory_overview'))
    products = conn.execute('SELECT id, style_no, name FROM products ORDER BY style_no').fetchall()
    warehouses = conn.execute('SELECT * FROM warehouses').fetchall()
    conn.close()
    return render_template('inventory_count.html', products=products, warehouses=warehouses)

# ═══════════════════════ 仓库 ═══════════════════════
@app.route('/warehouses')
def warehouse_list():
    conn = get_db()
    whs = conn.execute('''
        SELECT w.*, (SELECT COALESCE(SUM(quantity),0) FROM inventory WHERE warehouse_id=w.id) as total_items
        FROM warehouses w ORDER BY w.id
    ''').fetchall()
    conn.close()
    return render_template('warehouse_list.html', warehouses=whs)

@app.route('/warehouses/add', methods=['POST'])
def warehouse_add():
    name = request.form['name'].strip()
    if not name:
        flash('仓库名不能为空', 'error')
        return redirect(url_for('warehouse_list'))
    conn = get_db()
    try:
        conn.execute('INSERT INTO warehouses (name) VALUES (?)', (name,))
        conn.commit()
        flash(f'仓库 {name} 已添加', 'success')
    except sqlite3.IntegrityError:
        flash('仓库名已存在', 'error')
    conn.close()
    return redirect(url_for('warehouse_list'))

# ═══════════════════════ 财务 ═══════════════════════
@app.route('/finance')
def finance():
    conn = get_db()
    month = request.args.get('month', date.today().strftime('%Y-%m'))
    expenses = conn.execute('''
        SELECT category, COALESCE(SUM(amount),0) as total FROM expenses
        WHERE strftime('%Y-%m',created_at)=? GROUP BY category ORDER BY total DESC
    ''', (month,)).fetchall()
    # 月度汇总
    total_income = conn.execute("""
        SELECT COALESCE(SUM(final_amount),0) as amt FROM sales_orders
        WHERE strftime('%Y-%m',created_at)=? AND status='已完成'
    """, (month,)).fetchone()['amt']
    total_expense = conn.execute("""
        SELECT COALESCE(SUM(amount),0) as amt FROM expenses WHERE strftime('%Y-%m',created_at)=?
    """, (month,)).fetchone()['amt']
    total_purchase = conn.execute("""
        SELECT COALESCE(SUM(total_amount),0) as amt FROM purchase_orders
        WHERE strftime('%Y-%m',created_at)=?
    """, (month,)).fetchone()['amt']
    recent = conn.execute("""
        SELECT '收入' as type, final_amount as amount, order_no as ref, created_at FROM sales_orders
        WHERE strftime('%Y-%m',created_at)=? AND status='已完成'
        UNION ALL
        SELECT '支出' as type, amount, description as ref, created_at FROM expenses
        WHERE strftime('%Y-%m',created_at)=?
        ORDER BY created_at DESC LIMIT 100
    """, (month, month)).fetchall()
    conn.close()
    return render_template('finance.html', month=month, expenses=expenses,
                          total_income=total_income, total_expense=total_expense,
                          total_purchase=total_purchase, recent=recent)

@app.route('/finance/expense/add', methods=['POST'])
def expense_add():
    category = request.form.get('category','其他').strip()
    amount = float(request.form.get('amount',0))
    description = request.form.get('description','').strip()
    if amount <= 0:
        flash('金额必须大于0', 'error')
        return redirect(url_for('finance'))
    conn = get_db()
    conn.execute('INSERT INTO expenses (category,amount,description) VALUES (?,?,?)',
                (category,amount,description))
    conn.commit()
    conn.close()
    flash(f'支出已记录 ¥{amount:.2f}', 'success')
    return redirect(url_for('finance'))

# ═══════════════════════ 报表 ═══════════════════════
@app.route('/reports/sales')
def report_sales():
    conn = get_db()
    period = request.args.get('period','month')
    now = datetime.now()
    if period == 'today':
        gte = now.strftime('%Y-%m-%d')
        fmt = '%H:00'
    elif period == 'week':
        gte = (now - timedelta(days=now.weekday())).strftime('%Y-%m-%d')
        fmt = '%m-%d'
    elif period == 'year':
        gte = now.strftime('%Y-01-01')
        fmt = '%m'
    else:
        gte = now.strftime('%Y-%m-01')
        fmt = '%m-%d'
    sql = """
        SELECT strftime(?, created_at) as label,
               COUNT(*) as order_count,
               COALESCE(SUM(final_amount),0) as amount,
               COALESCE(SUM(final_amount - si.cost_price * si.quantity),0) as profit
        FROM sales_orders so
        JOIN sales_items si ON si.order_id=so.id
        WHERE created_at>=? AND status='已完成'
        GROUP BY label ORDER BY label
    """
    rows = conn.execute(sql, (fmt, gte)).fetchall()
    # 排行榜
    top_products = conn.execute("""
        SELECT p.style_no, p.name, p.color,
               SUM(si.quantity) as qty, SUM(si.subtotal) as amount
        FROM sales_items si
        JOIN products p ON p.id=si.product_id
        JOIN sales_orders so ON so.id=si.order_id
        WHERE so.created_at>=? AND so.status='已完成'
        GROUP BY si.product_id
        ORDER BY qty DESC LIMIT 20
    """, (gte,)).fetchall()
    top_customers = conn.execute("""
        SELECT c.name, COUNT(*) as order_count, COALESCE(SUM(so.final_amount),0) as amount
        FROM sales_orders so JOIN customers c ON c.id=so.customer_id
        WHERE so.created_at>=? AND so.status='已完成'
        GROUP BY so.customer_id ORDER BY amount DESC LIMIT 10
    """, (gte,)).fetchall()
    conn.close()
    return render_template('report_sales.html', rows=rows, period=period,
                          top_products=top_products, top_customers=top_customers)

@app.route('/reports/profit')
def report_profit():
    conn = get_db()
    rows = conn.execute("""
        SELECT strftime('%Y-%m',so.created_at) as month,
               COUNT(DISTINCT so.id) as orders,
               COALESCE(SUM(so.final_amount),0) as revenue,
               COALESCE(SUM(si.cost_price * si.quantity),0) as cost,
               COALESCE(SUM(so.final_amount - si.cost_price * si.quantity),0) as profit
        FROM sales_orders so
        JOIN sales_items si ON si.order_id=so.id
        WHERE so.status='已完成'
        GROUP BY month ORDER BY month
    """).fetchall()
    conn.close()
    return render_template('report_profit.html', rows=rows)

@app.route('/reports/inventory')
def report_inventory():
    conn = get_db()
    # 库存价值
    total_cost_value = conn.execute("""
        SELECT COALESCE(SUM(i.quantity * p.cost_price),0) as v
        FROM inventory i JOIN products p ON p.id=i.product_id
    """).fetchone()['v']
    total_retail_value = conn.execute("""
        SELECT COALESCE(SUM(i.quantity * p.retail_price),0) as v
        FROM inventory i JOIN products p ON p.id=i.product_id
    """).fetchone()['v']
    products = conn.execute("""
        SELECT p.style_no, p.name, p.color, p.cost_price, p.retail_price,
               COALESCE(SUM(i.quantity),0) as qty,
               COALESCE(SUM(i.quantity * p.cost_price),0) as cost_value,
               COALESCE(SUM(i.quantity * p.retail_price),0) as retail_value
        FROM products p
        LEFT JOIN inventory i ON i.product_id=p.id
        GROUP BY p.id HAVING qty>0
        ORDER BY cost_value DESC
    """).fetchall()
    conn.close()
    return render_template('report_inventory.html', total_cost_value=total_cost_value,
                          total_retail_value=total_retail_value, products=products)

# ═══════════════════════ 导出 ═══════════════════════
@app.route('/export/<string:type>')
def export_excel(type):
    conn = get_db()
    fp = os.path.join(BASE_DIR, f'export_{type}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    wb = xlsxwriter.Workbook(fp)
    if type == 'inventory':
        ws = wb.add_worksheet('库存')
        headers = ['款号','款名','颜色','尺码','数量','成本价','零售价','成本金额','零售金额']
        ws.write_row(0,0,headers)
        rows = conn.execute('''
            SELECT p.style_no, p.name, p.color, i.size, i.quantity, p.cost_price, p.retail_price
            FROM inventory i JOIN products p ON p.id=i.product_id
            WHERE i.quantity>0 ORDER BY p.style_no, i.size
        ''').fetchall()
        for i, r in enumerate(rows, 1):
            ws.write_row(i,0,[r['style_no'],r['name'],r['color'],r['size'],r['quantity'],
                             r['cost_price'],r['retail_price'],
                             r['cost_price']*r['quantity'], r['retail_price']*r['quantity']])
    elif type == 'products':
        ws = wb.add_worksheet('商品')
        headers = ['款号','款名','颜色','条码','分类','品牌','进货价','零售价','批发价','总库存']
        ws.write_row(0,0,headers)
        rows = conn.execute('''
            SELECT p.*, (SELECT COALESCE(SUM(quantity),0) FROM inventory WHERE product_id=p.id) as stock
            FROM products p ORDER BY p.style_no
        ''').fetchall()
        for i, r in enumerate(rows, 1):
            ws.write_row(i,0,[r['style_no'],r['name'],r['color'],r['barcode'],
                             '', '', r['cost_price'],r['retail_price'],r['wholesale_price'],r['stock']])
    elif type == 'sales':
        ws = wb.add_worksheet('销售')
        headers = ['单号','日期','客户','金额','支付方式','商品明细']
        ws.write_row(0,0,headers)
        rows = conn.execute('''
            SELECT so.*, COALESCE(c.name,'') as cname FROM sales_orders so
            LEFT JOIN customers c ON c.id=so.customer_id
            ORDER BY so.created_at DESC LIMIT 1000
        ''').fetchall()
        for i, r in enumerate(rows, 1):
            items = conn.execute('''
                SELECT si.*, p.style_no FROM sales_items si JOIN products p ON p.id=si.product_id
                WHERE si.order_id=?
            ''', (r['id'],)).fetchall()
            detail = '; '.join([f"{it['style_no']} x{it['quantity']}" for it in items])
            ws.write_row(i,0,[r['order_no'],r['created_at'][:10],r['cname'],r['final_amount'],r['payment_method'],detail])
    wb.close()
    conn.close()
    return send_file(fp, as_attachment=True, download_name=f'{type}_{date.today().strftime("%Y%m%d")}.xlsx')

# ═══════════════════════ 启动 ═══════════════════════
if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
