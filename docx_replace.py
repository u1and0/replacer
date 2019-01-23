#!/usr/bin/env python
# -*- coding: utf-8 -*-
from docx import Document

document = Document("test/テスト.docx")

for paragraph in document.paragraphs:
    paragraph.text = paragraph.text.replace(u"テスト", u"置換")

document.save("test/rep_テスト.docx")
