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
            network_area TEXT DEFAULT '生产网'
        )
    ''')
    # Add network_area column if it doesn't exist (for existing tables)
    try:
        cursor.execute('ALTER TABLE accounts ADD COLUMN network_area TEXT DEFAULT "生产网"')
    except sqlite3.OperationalError:
        # Column already exists
        pass
    # Update existing records to set network_area to '生产网' if it's NULL
    cursor.execute('UPDATE accounts SET network_area = "生产网" WHERE network_area IS NULL')
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
        asset_number = request.form['asset_number']
        computer_name = request.form['computer_name']
        phone_number = request.form['phone_number']
        department = request.form['department']
        network_area = request.form['network_area'] or '生产网'
        
        if not user_name:
            flash('用户姓名是必填项')
            return redirect(url_for('add_account'))
        
        conn = get_db_connection()
        conn.execute('''
            INSERT INTO accounts (user_name, account_number, asset_number, computer_name, phone_number, department, network_area)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (user_name, account_number, asset_number, computer_name, phone_number, department, network_area))
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
        asset_number = request.form['asset_number']
        computer_name = request.form['computer_name']
        phone_number = request.form['phone_number']
        department = request.form['department']
        network_area = request.form['network_area'] or '生产网'
        
        if not user_name:
            flash('用户姓名是必填项')
            return render_template('edit.html', account=account)
        
        conn.execute('''
            UPDATE accounts
            SET user_name = ?, account_number = ?, asset_number = ?, computer_name = ?, phone_number = ?, department = ?, network_area = ?
            WHERE id = ?
        ''', (user_name, account_number, asset_number, computer_name, phone_number, department, network_area, id))
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
                    # Check if we have 7 columns (including network_area)
                    if len(df.columns) >= 7:
                        df.columns = ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department', 'network_area']
                    else:
                        df.columns = ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department']
                        df['network_area'] = '生产网'
                else:
                    # Map Chinese column names to database field names
                    column_mapping = {
                        '用户姓名': 'user_name',
                        '账号': 'account_number', 
                        '资产编号': 'asset_number',
                        '计算机名': 'computer_name',
                        '联系电话': 'phone_number',
                        '所在部门': 'department',
                        '网络区域': 'network_area'
                    }
                    df = df.rename(columns=column_mapping)
                    
                    # Ensure all required columns exist after mapping
                    for col in ['user_name', 'account_number', 'asset_number', 'computer_name', 'phone_number', 'department', 'network_area']:
                        if col not in df.columns:
                            df[col] = '' if col != 'network_area' else '生产网'  # Add missing columns with empty values, default network_area to '生产网'
                
                # Insert data into database
                conn = get_db_connection()
                for _, row in df.iterrows():
                    conn.execute('''
                        INSERT INTO accounts (user_name, account_number, asset_number, computer_name, phone_number, department, network_area)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        row.get('user_name', ''),
                        row.get('account_number', ''),
                        row.get('asset_number', ''),
                        row.get('computer_name', ''),
                        row.get('phone_number', ''),
                        row.get('department', ''),
                        row.get('network_area', '生产网')
                    ))
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
        'network_area': '网络区域'
    })
    # Remove the ID column for export as it's internal
    df = df[['用户姓名', '账号', '资产编号', '计算机名', '联系电话', '所在部门', '网络区域']]
    
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