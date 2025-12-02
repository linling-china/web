import os
import sqlite3
from flask import Flask, request, render_template, redirect, url_for, flash, send_file
import pandas as pd
from werkzeug.utils import secure_filename
import io

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Change this to a random secret key

# Database configuration
DATABASE = 'accounts.db'
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create uploads directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def init_db():
    """Initialize the database with accounts table"""
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    # Create table if it doesn't exist
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_name TEXT NOT NULL,
            account_number TEXT,
            asset_number TEXT,
            computer_name TEXT,
            phone_number TEXT,
            department TEXT,
            network_area TEXT DEFAULT '生产网',
            account_status TEXT DEFAULT '在用'
        )
    ''')
    # Add network_area column if it doesn't exist (for existing tables)
    try:
        cursor.execute('ALTER TABLE accounts ADD COLUMN network_area TEXT DEFAULT "生产网"')
    except sqlite3.OperationalError:
        # Column already exists
        pass
    # Add account_status column if it doesn't exist (for existing tables)
    try:
        cursor.execute('ALTER TABLE accounts ADD COLUMN account_status TEXT DEFAULT "在用"')
    except sqlite3.OperationalError:
        # Column already exists
        pass
    # Update existing records to set network_area to '生产网' if it's NULL
    cursor.execute('UPDATE accounts SET network_area = "生产网" WHERE network_area IS NULL')
    # Update existing records to set account_status to '在用' if it's NULL
    cursor.execute('UPDATE accounts SET account_status = "在用" WHERE account_status IS NULL')
    conn.commit()
    conn.close()

def get_db_connection():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row  # This enables column access by name
    return conn

def allowed_file(filename):
    """Check if file extension is allowed"""
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS

def get_prefixes_by_network_area(network_area):
    """Get prefixes based on network area"""
    prefixes = {
        '管理网': {
            'account_number': 'foc-',
            'asset_number': 'g-fj-foc-',
            'computer_name': 'G-FJ-FOC-'
        },
        '生产网': {
            'account_number': 'foc-d-',
            'asset_number': 'z-fj-foc-',
            'computer_name': 'Z-FJ-FOC-'
        },
        '金融网': {
            'account_number': 'foc-d-',
            'asset_number': 'l-fj-foc-',
            'computer_name': 'L-FJ-FOC-'
        }
    }
    return prefixes.get(network_area, prefixes['生产网'])

def add_prefix(value, prefix):
    """Add prefix to value if it doesn't already have it"""
    if not value:
        return ''
    if value.startswith(prefix):
        return value
    return f"{prefix}{value}"

def remove_prefix(value, prefixes):
    """Remove any of the given prefixes from value"""
    if not value:
        return ''
    for prefix in prefixes:
        if value.startswith(prefix):
            return value[len(prefix):]
    return value

@app.route('/')
def index():
    """Main page - display all accounts"""
    conn = get_db_connection()
    accounts = conn.execute('SELECT * FROM accounts ORDER BY id').fetchall()
    conn.close()
    return render_template('index.html', accounts=accounts)

@app.route('/add', methods=['GET', 'POST'])
def add_account():
    """Add a new account"""
    if request.method == 'POST':
        user_name = request.form['user_name']
        account_number = request.form['account_number']
        # asset_number is now auto-generated, ignore user input
        computer_name = request.form['computer_name']
        phone_number = request.form['phone_number']
        department = request.form['department']
        network_area = request.form['network_area'] or '生产网'
        account_status = request.form['account_status'] or '在用'
        
        if not user_name:
            flash('用户姓名是必填项')
            return redirect(url_for('add_account'))
        
        # Get prefixes based on network area
        prefixes = get_prefixes_by_network_area(network_area)
        
        # Add prefixes to fields (except asset_number which will be auto-generated)
        account_number = add_prefix(account_number, prefixes['account_number'])
        computer_name = add_prefix(computer_name, prefixes['computer_name'])
        
        conn = get_db_connection()
        # Insert record without asset_number (will be auto-generated later)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO accounts (user_name, account_number, computer_name, phone_number, department, network_area, account_status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (user_name, account_number, computer_name, phone_number, department, network_area, account_status))
        
        # Get the generated id
        account_id = cursor.lastrowid
        
        # Auto-generate asset_number: prefix + id
        asset_number = f"{prefixes['asset_number']}{account_id}"
        
        # Update the record with the generated asset_number
        cursor.execute('''
            UPDATE accounts SET asset_number = ? WHERE id = ?
        ''', (asset_number, account_id))
        
        conn.commit()
        conn.close()
        
        flash('账户信息已添加')
        return redirect(url_for('index'))
    
    return render_template('add.html')

@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit_account(id):
    """Edit an existing account"""
    conn = get_db_connection()
    account = conn.execute('SELECT * FROM accounts WHERE id = ?', (id,)).fetchone()
    
    if request.method == 'POST':
        user_name = request.form['user_name']
        account_number = request.form['account_number']
        # asset_number is auto-generated, don't modify it
        computer_name = request.form['computer_name']
        phone_number = request.form['phone_number']
        department = request.form['department']
        network_area = request.form['network_area'] or '生产网'
        account_status = request.form['account_status'] or '在用'
        
        if not user_name:
            flash('用户姓名是必填项')
            return render_template('edit.html', account=account)
        
        # Get all possible prefixes for removal
        all_account_prefixes = ['foc-', 'foc-d-']
        all_computer_prefixes = ['G-FJ-FOC-', 'Z-FJ-FOC-', 'L-FJ-FOC-']
        
        # Get new prefixes based on network area
        prefixes = get_prefixes_by_network_area(network_area)
        
        # Remove old prefixes and add new ones
        account_number = remove_prefix(account_number, all_account_prefixes)
        account_number = add_prefix(account_number, prefixes['account_number'])
        
        # asset_number is auto-generated, don't modify it
        
        computer_name = remove_prefix(computer_name, all_computer_prefixes)
        computer_name = add_prefix(computer_name, prefixes['computer_name'])
        
        conn.execute('''
            UPDATE accounts
            SET user_name = ?, account_number = ?, computer_name = ?, phone_number = ?, department = ?, network_area = ?, account_status = ?
            WHERE id = ?
        ''', (user_name, account_number, computer_name, phone_number, department, network_area, account_status, id))
        conn.commit()
        conn.close()
        
        flash('账户信息已更新')
        return redirect(url_for('index'))
    
    return render_template('edit.html', account=account)

@app.route('/delete/<int:id>', methods=['POST'])
def delete_account(id):
    """Delete an account"""
    conn = get_db_connection()
    conn.execute('DELETE FROM accounts WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    
    flash('账户信息已删除')
    return redirect(url_for('index'))

@app.route('/import', methods=['GET', 'POST'])
def import_excel():
    """Import accounts from Excel file"""
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('没有选择文件')
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            flash('没有选择文件')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            try:
                # Read the Excel file into a pandas DataFrame
                df = pd.read_excel(file)
                
                # Ensure required columns exist
                required_columns = ['用户姓名', '账号', '资产编号', '计算机名', '联系电话', '所在部门']
                
                # If the Excel doesn't have the expected column names, we'll map them by position
                # Or we can check if these exact column names exist
                if '用户姓名' not in df.columns:
                    # If standard names don't exist, assume they are in order: user_name, account_number, etc.
                    # Check if we have 8 columns (including network_area and account_status)
                    if len(df.columns) >= 8:
                        df.columns = ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department', 'network_area', 'account_status']
                    elif len(df.columns) >= 7:
                        df.columns = ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department', 'network_area']
                        df['account_status'] = '在用'
                    else:
                        df.columns = ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department']
                        df['network_area'] = '生产网'
                        df['account_status'] = '在用'
                else:
                    # Map Chinese column names to database field names
                    column_mapping = {
                        '用户姓名': 'user_name',
                        '账号': 'account_number', 
                        '资产编号': 'asset_number',
                        '计算机名': 'computer_name',
                        '联系电话': 'phone_number',
                        '所在部门': 'department',
                        '网络区域': 'network_area',
                        '账号状态': 'account_status'
                    }
                    df = df.rename(columns=column_mapping)
                    
                    # Ensure all required columns exist after mapping
                    for col in ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department', 'network_area', 'account_status']:
                        if col not in df.columns:
                            if col == 'network_area':
                                df[col] = '生产网'
                            elif col == 'account_status':
                                df[col] = '在用'
                            else:
                                df[col] = ''
                
                # Get all possible prefixes for removal
                all_account_prefixes = ['foc-', 'foc-d-']
                all_computer_prefixes = ['G-FJ-FOC-', 'Z-FJ-FOC-', 'L-FJ-FOC-']
                
                # Insert data into database
                conn = get_db_connection()
                cursor = conn.cursor()
                
                for _, row in df.iterrows():
                    user_name = row.get('user_name', '')
                    account_number = row.get('account_number', '')
                    # asset_number is now auto-generated, ignore user input
                    computer_name = row.get('computer_name', '')
                    phone_number = row.get('phone_number', '')
                    department = row.get('department', '')
                    network_area = row.get('network_area', '生产网')
                    account_status = row.get('account_status', '在用')
                    
                    # Get prefixes based on network area
                    prefixes = get_prefixes_by_network_area(network_area)
                    
                    # Remove old prefixes and add new ones
                    account_number = remove_prefix(account_number, all_account_prefixes)
                    account_number = add_prefix(account_number, prefixes['account_number'])
                    
                    # asset_number is auto-generated, don't modify it here
                    
                    computer_name = remove_prefix(computer_name, all_computer_prefixes)
                    computer_name = add_prefix(computer_name, prefixes['computer_name'])
                    
                    # Insert record without asset_number (will be auto-generated later)
                    cursor.execute('''
                        INSERT INTO accounts (user_name, account_number, computer_name, phone_number, department, network_area, account_status)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        user_name,
                        account_number,
                        computer_name,
                        phone_number,
                        department,
                        network_area,
                        account_status
                    ))
                    
                    # Get the generated id
                    account_id = cursor.lastrowid
                    
                    # Auto-generate asset_number: prefix + id
                    asset_number = f"{prefixes['asset_number']}{account_id}"
                    
                    # Update the record with the generated asset_number
                    cursor.execute('''
                        UPDATE accounts SET asset_number = ? WHERE id = ?
                    ''', (asset_number, account_id))
                
                conn.commit()
                conn.close()
                
                flash(f'成功导入 {len(df)} 条记录')
                return redirect(url_for('index'))
                
            except Exception as e:
                flash(f'导入失败: {str(e)}')
                return redirect(url_for('import_excel'))
        else:
            flash('不支持的文件格式，请上传Excel文件(.xlsx或.xls)')
            return redirect(url_for('import_excel'))
    
    return render_template('import.html')

@app.route('/export', endpoint='export')
def export_excel():
    """Export accounts to Excel file"""
    conn = get_db_connection()
    accounts = conn.execute('SELECT * FROM accounts ORDER BY id').fetchall()
    conn.close()
    
    # Convert to DataFrame
    df = pd.DataFrame(accounts)
    # Rename columns to Chinese names for the export
    df = df.rename(columns={
        'id': 'ID',
        'user_name': '用户姓名',
        'account_number': '账号', 
        'asset_number': '资产编号',
        'computer_name': '计算机名',
        'phone_number': '联系电话',
        'department': '所在部门',
        'network_area': '网络区域',
        'account_status': '账号状态'
    })
    # Remove the ID column for export as it's internal
    df = df[['用户姓名', '账号', '资产编号', '计算机名', '联系电话', '所在部门', '网络区域', '账号状态']]
    
    # Create a BytesIO buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='账户信息')
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='accounts_export.xlsx'
    )

if __name__ == '__main__':
    init_db()
    app.run(debug=True, host='0.0.0.0', port=5000)