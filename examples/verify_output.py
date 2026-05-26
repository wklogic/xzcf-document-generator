#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证生成的文档"""

from pathlib import Path
from docx import Document

output_path = Path(__file__).parent / "输出文件"
files = sorted(output_path.glob("*.docx"))
if files:
    doc = Document(str(files[0]))
else:
    print("输出文件目录下未找到 .docx 文件")
    exit(1)

print("生成的文档内容预览：")
print("=" * 50)
for i, para in enumerate(doc.paragraphs[:15]):
    if para.text.strip():
        print(para.text)
print("=" * 50)
