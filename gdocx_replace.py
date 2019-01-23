#!/usr/bin/env python
import sys
import os
import tkinter
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from docx import Document

root = tkinter.Tk()
root.title(u"docx置換ツール")
root.geometry("300x120")


def input_value(event):
    fTyp = [('wordファイル', "*.docx")]
    iDir = os.getcwd()
    filename = filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
    before = EditBox_before.get()
    after = EditBox_after.get()
    print(before, after)

    document = Document(filename)
    for para in document.paragraphs:
        para.text = para.text.replace(before, after)

    document.save(filename.replace(u".docx", "_replace.docx"))

    messagebox.showinfo('変換完了')


# 置換前
# ラベル
Static1 = tkinter.Label(text=u'置換前テキスト')
Static1.place(x=0, y=10)
# エントリー
EditBox_before = tkinter.Entry(width=20)
# EditBox.insert(tkinter.END,"挿入する文字列")
EditBox_before.place(x=100, y=10)

# 置換後
# ラベル
Static2 = tkinter.Label(text=u'置換後テキスト')
Static2.place(x=0, y=40)
# エントリー
EditBox_after = tkinter.Entry(width=20)
# EditBox.insert(tkinter.END,"挿入する文字列")
EditBox_after.place(x=100, y=40)

# ボタン
Button = ttk.Button(text=u'docxファイル選択＆変換', width=20)
Button.bind("<Button-1>", input_value)
Button.place(x=20, y=70)

root.mainloop()
