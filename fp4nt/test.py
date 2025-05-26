from fp4nt import Fp4NT

fp4nt = Fp4NT(row_fill_merged=True, col_fill_merged=False)

# 示例1：HTML表格转Markdown
html = "<table><tr><th>Header1</th><th>Header2</th></tr><tr><td>Data1</td><td>Data2</td></tr></table>"
markdown_output = fp4nt.convert_html_to_markdown(html)
print("HTML转Markdown结果：")
print(markdown_output)

# 示例2：Excel转DataFrame
data_path = '/data/yfzx/lrs/fp4nt/test.xlsx'
df = fp4nt.build_excel_frame(data_path)
print("\nExcel转DataFrame结果：")
print(df.to_markdown(index=False))