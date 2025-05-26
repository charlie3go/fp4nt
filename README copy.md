# 📚 fp4nt —— 嵌套表格解析工具
---

![License](https://img.shields.io/pypi/l/fp4nt)

> 🧰 专为嵌套表格数据设计的解析与转换工具，支持 HTML 和 Excel 转 Markdown 表格

---

## 🌟 功能亮点

- ✅ 支持从 **HTML** 提取并转换为 **Markdown 表格**
- ✅ 支持处理 **Excel（.xlsx）** 文件中的合并单元格
- ✅ 自动填充 **跨行/跨列空白单元格**
- ✅ 可自定义 **填充标记**（如“无”）
- ✅ 简洁的日志系统便于调试

---

## 🛠️ 安装方式

```bash
pip install fp4nt
```

或升级至最新版本：

```bash
pip install --upgrade fp4nt
```

---

## 🔧 使用方法

### 1. 初始化对象

```python
from fp4nt import Fp4NT

# 可选参数控制是否填充合并单元格
fp4nt = Fp4NT(row_fill_merged=True, col_fill_merged=False, fill_mark='无')
```

### 2. HTML 表格转 Markdown

```python
html = "<table><tr><th>Header1</th><th>Header2</th></tr><tr><td>Data1</td><td>Data2</td></tr></table>"
markdown_output = fp4nt.convert_html_to_markdown(html)
print(markdown_output)
```

### 3. Excel 重构 DataFrame（含合并单元格处理）

```python
data_path = '/path/to/your/file.xlsx'
df = fp4nt.build_excel_frame(data_path)
print(df.to_markdown(index=False))
```

---

## 📦 依赖库

确保已安装以下依赖包：

```bash
pip install -r requirements.txt
```

依赖项如下：

| 包名             | 版本要求       |
|------------------|----------------|
| beautifulsoup4   | >=4.9.0        |
| html2text        | >=2020.1.16    |
| pandas           | >=1.0.0        |
| openpyxl         | >=3.0.0        |

---

## 📁 示例文件结构

```
fp4nt/
├── __init__.py
├── fp4nt.py      # 主模块
└── README.md     # 当前文档
```

---

## 📄 示例输出（Markdown 表格）

输入 HTML：

```html
<table>
  <tr><th>姓名</th><th>年龄</th></tr>
  <tr><td>张三</td><td>25</td></tr>
</table>
```

输出 Markdown：

```
| 姓名 | 年龄 |
|------|------|
| 张三 | 25   |
```

---

## 💬 反馈建议 & 报告问题

欢迎在 [GitHub Issues](https://github.com/yourname/fp4nt/issues) 中提交反馈或 bug 报告！

---

## 📜 协议

该项目采用 MIT License，请参阅 [LICENSE](LICENSE) 文件了解更多。

---

## 🙌 感谢

感谢你使用 `fp4nt`！如果你觉得这个工具对你有帮助，欢迎 ⭐ Star、分享给朋友或贡献代码 😊

---