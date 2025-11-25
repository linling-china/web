# 账户信息管理系统

这是一个基于Flask的Web应用程序，用于管理账户信息，支持Excel文件的导入和导出功能。

## 功能特性

- **查看账户信息**：以表格形式展示所有账户信息
- **添加账户**：通过Web表单添加新账户
- **编辑账户**：修改现有账户信息
- **删除账户**：删除不需要的账户
- **导入Excel**：从Excel文件批量导入账户信息
- **导出Excel**：将账户信息导出为Excel文件
- **搜索功能**：在页面上搜索特定账户信息

## 技术栈

- **后端框架**：Flask
- **数据库**：SQLite
- **数据处理**：Pandas
- **前端**：HTML, CSS, JavaScript

## 安装和运行

### 环境要求

- Python 3.6+
- pip

### 安装依赖

```bash
pip install Flask pandas openpyxl
```

### 运行应用

```bash
python app.py
```

应用将在 `http://localhost:5000` 上运行。

## Excel文件格式

导入的Excel文件应包含以下列（列名可以是中文或按顺序排列）：

1. 用户姓名（必填）
2. 账号
3. 资产编号
4. 计算机名
5. 联系电话
6. 所在部门

## 使用说明

1. **添加账户**：点击"添加新账户"按钮，填写表单后提交
2. **编辑账户**：在账户列表中点击"编辑"按钮进行修改
3. **删除账户**：在账户列表中点击"删除"按钮进行删除
4. **导入Excel**：点击"导入Excel"按钮，选择Excel文件进行批量导入
5. **导出Excel**：点击"导出Excel"按钮，下载当前所有账户信息

## 数据库结构

应用使用SQLite数据库，创建名为`accounts.db`的数据库文件，包含以下表结构：

- `accounts` 表：
  - `id`: 主键，自增整数
  - `user_name`: 用户姓名（必填）
  - `account_number`: 账号
  - `asset_number`: 资产编号
  - `computer_name`: 计算机名
  - `phone_number`: 联系电话
  - `department`: 所在部门

## 多用户支持

应用支持多人同时进行新增操作，使用SQLite的并发处理机制确保数据一致性。

## 文件结构

```
/workspace/
├── app.py                 # 主应用文件
├── accounts.db           # SQLite数据库文件
├── uploads/              # 上传文件目录
├── templates/            # HTML模板目录
│   ├── index.html        # 主页面
│   ├── add.html          # 添加账户页面
│   ├── edit.html         # 编辑账户页面
│   └── import.html       # 导入Excel页面
├── sample_accounts.xlsx  # 示例Excel文件
├── requirements.txt      # 依赖包列表
└── README.md             # 说明文档
```

## 注意事项

- 应用使用开发服务器，不适用于生产环境
- 数据库文件为`accounts.db`，请妥善备份
- Excel导入支持`.xlsx`和`.xls`格式
- 用户姓名为必填字段