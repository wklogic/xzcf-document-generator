# 通用模板渲染工具

基于 Python 的批量文档生成工具——用 Word 模板 + Excel 数据，批量生成 `docx` 文件。

## 功能特点

- **占位符模板**：使用 `{{字段名}}` 语法，肉眼可见，不易出错
- **两种数据布局**：水平布局（宽表）和垂直布局（每记录一个Sheet），自动识别
- **配置驱动**：通过 JSON 配置文件管理所有参数，无需修改代码
- **系统变量注入**：自动填充日期、时间等动态值
- **灵活命名**：输出文件名由模板控制，支持数据字段和系统变量
- **详细日志**：同时输出到文件和控制台，便于排查问题
- **跨平台**：纯 Python，支持 Windows、Linux、macOS

## 目录结构

```
├── render.py                 # 主程序
├── config.json               # 配置文件
├── templates/                # Word 模板目录
│   └── template.docx
├── data/                     # Excel 数据目录
│   └── data.xlsx
├── output/                   # 输出目录
├── examples/                 # 示例脚本和示例数据
│   ├── create_template.py
│   ├── create_sample_data_horizontal.py
│   ├── create_sample_data_vertical.py
│   └── verify_output.py
├── 开始生成.bat              # 一键启动
├── 安装依赖.bat              # 安装依赖
└── build.bat                 # 打包为 exe
```

## 快速开始

### 1. 安装依赖

```bash
pip install docxtpl pandas openpyxl python-docx
```

### 2. 准备模板

在 `templates/` 目录下创建 Word 模板，使用 `{{字段名}}` 作为占位符：

```
关于{{案件名称}}的处理决定

文号：{{文号}}

当事人：{{当事人名称}}
事由：{{案由}}

处理决定：{{处罚决定}}

签发日期：{{render_date}}
```

### 3. 准备数据

支持两种数据布局，程序自动识别：

#### 水平布局（单Sheet宽表）

| 文号 | 当事人名称 | 案件名称 | 案由 | 处罚决定 |
|------|-----------|---------|------|---------|
| 2024-001 | 张三 | XX案 | XX | 罚款5万元 |

#### 垂直布局（每记录一个Sheet，两列：字段名 | 值）

**Sheet: 2024-001**

| 字段名 | 值 |
|--------|-----|
| 文号 | 2024-001 |
| 当事人名称 | 张三 |
| ... | ... |

### 4. 运行

```bash
python render.py
```

## 配置说明

编辑 `config.json`：

```json
{
  "template_file": "templates/template.docx",
  "data_file": "data/data.xlsx",
  "output_dir": "output",
  "filename_template": "{案号}_{当事人名称}.docx",
  "sheet_name": 0,
  "skip_rows": 0,
  "skip_sheets": [],
  "layout": "auto",
  "system_variables": {
    "render_date": "%Y年%m月%d日",
    "render_time": "%Y%m%d_%H%M%S",
    "render_time_readable": "%Y-%m-%d %H:%M:%S",
    "render_year": "%Y"
  }
}
```

| 配置项 | 说明 | 示例 |
|--------|------|------|
| `template_file` | 模板文件路径 | `templates/template.docx` |
| `data_file` | 数据文件路径 | `data/data.xlsx` |
| `output_dir` | 输出目录 | `output` |
| `filename_template` | 文件名模板，支持 `{字段名}` 和系统变量 | `{案号}_{当事人名称}.docx` |
| `sheet_name` | 水平布局时的Sheet索引或名称 | `0` 或 `Sheet1` |
| `skip_rows` | 水平布局时跳过顶部行数 | `0` |
| `skip_sheets` | 垂直布局时跳过的Sheet名称列表 | `["说明", "汇总"]` |
| `layout` | 布局模式：`auto`（自动）、`horizontal`、`vertical` | `auto` |
| `system_variables` | 自定义系统变量，键为变量名，值为 `strftime` 格式 | 见上方 |

## 系统变量

模板中可以使用以下系统变量（无需在Excel中提供）：

| 变量 | 格式示例 | 说明 |
|------|---------|------|
| `{{render_date}}` | 2024年01月15日 | 生成日期 |
| `{{render_time}}` | 20240115_143000 | 紧凑时间（适合文件名） |
| `{{render_time_readable}}` | 2024-01-15 14:30:00 | 可读时间（适合正文） |
| `{{render_year}}` | 2024 | 当前年份 |

兼容旧版中文变量名：`{{生成日期}}`、`{{生成时间}}`、`{{生成时间_可读}}`、`{{年份}}` 仍然可用。

用户可在 `config.json` 的 `system_variables` 中自定义（增删改）系统变量。

## 占位符语法（Jinja2）

### 基本用法

```
当事人：{{当事人名称}}
处理结果：{{处理结果}}
```

### 条件判断

```
{% if 性别 == '男' %}先生{% else %}女士{% endif %}
```

### 循环

```
{% for item in 列表字段 %}
{{ loop.index }}. {{ item }}
{% endfor %}
```

## 打包为 EXE

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包
pyinstaller -F --clean render.py --name render --distpath . -y
```

或直接运行 `build.bat`。

## 常见问题

**Q: 模板中的占位符没有被替换？**

检查：1. 占位符格式是否为 `{{字段名}}`（双大括号）；2. Excel 字段名是否与占位符完全一致；3. 查看 `render.log`。

**Q: 如何添加新字段？**

在 Excel 中加新列，在模板中加 `{{新字段名}}`，无需改代码。

**Q: 支持什么格式？**

模板仅 `.docx`；数据支持 `.xlsx` 和 `.xls`。

## 许可证

MIT License
