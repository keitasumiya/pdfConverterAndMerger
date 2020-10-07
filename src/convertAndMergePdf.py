#!python3
# -*- coding:utf-8 -*-

import PyPDF2
from PyPDF2 import PdfFileMerger
import glob
import os, tkinter, tkinter.filedialog, tkinter.messagebox
import shutil
import traceback
import doc2pdf
import xls2pdf
import ppt2pdf
import img2pdfConverter


def moveFileInto(_path, _pathMoved, _ext) :
    filelistTmp = glob.glob(_path+'*.'+_ext)
    for file in filelistTmp: 
        shutil.move(file, _pathMoved)

def moveFilesInto(_path, _pathMoved) :
    moveFileInto(path, pathMoved, 'png')
    moveFileInto(path, pathMoved, 'jpeg')
    moveFileInto(path, pathMoved, 'jpg')
    moveFileInto(path, pathMoved, 'doc')
    moveFileInto(path, pathMoved, 'docx')
    moveFileInto(path, pathMoved, 'docm')
    moveFileInto(path, pathMoved, 'ppt')
    moveFileInto(path, pathMoved, 'pptx')
    moveFileInto(path, pathMoved, 'pptm')
    moveFileInto(path, pathMoved, 'pps')
    moveFileInto(path, pathMoved, 'ppsx')
    moveFileInto(path, pathMoved, 'ppsm')
    moveFileInto(path, pathMoved, 'xls')
    moveFileInto(path, pathMoved, 'xlsx')
    moveFileInto(path, pathMoved, 'xlsm')


# フォルダ選択ダイアログの表示
root = tkinter.Tk()
root.withdraw()
fTyp = [("","*")]
iDir = os.path.abspath(os.path.dirname(__file__))
tkinter.messagebox.showinfo('○×プログラム','整理したい対象フォルダを今から選択してください')

dir = tkinter.filedialog.askdirectory(initialdir = iDir)

print(dir)

doc2pdf.doc2pdfIn(dir)
ppt2pdf.ppt2pdfIn(dir)
xls2pdf.xls2pdfIn(dir)
img2pdfConverter.findFilesAndConvertToPdf(dir, '*.png')
img2pdfConverter.findFilesAndConvertToPdf(dir, '*.jpeg')
img2pdfConverter.findFilesAndConvertToPdf(dir, '*.jpg')

path = dir + '\\'
print(path)

pathMoved = path + '結合前ファイル/'
os.makedirs(pathMoved, exist_ok=True)

moveFilesInto(path, pathMoved)

filelist = glob.glob(path+'*.pdf')

submits = []
for file in filelist: 
    print('hello')
    print(file)
    submit = file.rsplit('_', 1)[0]
    print(submit)
    print(repr(submit))
    submits.append(submit)

print("")

submitsUnique = list(set(submits))

for submit in submitsUnique: 
    print(submit)
    num = submits.count(submit)
    print("file num = "+str(num))
    if num > 1 :
        try:
            writer = PyPDF2.PdfFileWriter()
            sourcePdfs = []
            for i in range(1, num+1, 1):
                print("use pdf-reader"+str(i))
                sourcePdfs.append(open(submit+"_"+str(i)+'.pdf', 'rb'))
                pages = PyPDF2.PdfFileReader(sourcePdfs[i-1], strict=False)
                print("num of pages = "+str(pages.getNumPages()) )
                for page in range(pages.getNumPages()):
                    writer.addPage(pages.getPage(page))

            with open(submit+'.pdf', 'wb') as f:
                writer.write(f)

            for i in range(num):
                sourcePdfs[i].close()

            for i in range(1, num+1, 1):
                shutil.move(submit+"_"+str(i)+'.pdf', pathMoved)
        except PyPDF2.utils.PdfReadError as e:
            with open(path+'_error.txt', 'a', encoding='utf-8') as f:
                print(submit, file=f)
            print(f"ERROR: {e} PDFが正しくありません")
            traceback.print_exc()
        except Exception as e:
            with open(path+'_error.txt', 'a', encoding='utf-8') as f:
                print(submit, file=f)
            print(f"ERROR: {e} 不明なエラー")
            traceback.print_exc()



    if submit == submitsUnique[-1] :
        tkinter.messagebox.showinfo('○×プログラム','整理が終わりました')

