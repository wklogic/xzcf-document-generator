#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证生成的文档"""

from docx import Document

doc = Document('G:/XZCF/行政处罚生成工具/输出文件/2024-001_张三_行政处罚决定书.docx')

print("生成的文档内容预览：")
print("=" * 50)
for i, para in enumerate(doc.paragraphs[:15]):
    if para.text.strip():
        print(para.text)
print("=" * 50)
