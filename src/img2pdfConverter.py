import os
import img2pdf
from PIL import Image
import glob
import tkinter, tkinter.filedialog, tkinter.messagebox


def convertImageToPdf (_filename):
    dirname = os.path.dirname(_filename)
    fileBasename = os.path.basename(_filename)
    filenameWithoutExt = os.path.splitext(fileBasename)[0]
    fileExt = os.path.splitext(fileBasename)[1]
    outputFilename = dirname + "/" + filenameWithoutExt + ".pdf"

    with open(outputFilename, "wb") as f:
        inputFileName = dirname + "/" + fileBasename
        if fileExt == ".png" or fileExt == ".PNG" :
            midFileName = dirname + "/" + filenameWithoutExt + ".jpg"
            im = Image.open(inputFileName)
            im = im.convert("RGB")
            im.save(midFileName, quality=100)
            f.write(img2pdf.convert(midFileName))
            os.remove(midFileName)
        else : 
            f.write(img2pdf.convert(inputFileName))

def findFilesAndConvertToPdf(_target_path, _ext) :
    print('converting '+_ext+' files ...')
    filelist = glob.glob(_target_path + '\\' + _ext)
    for file in filelist: 
        print(file)
        filename = file
        convertImageToPdf(filename)


if __name__ == '__main__':
    # フォルダ選択ダイアログの表示
    root = tkinter.Tk()
    root.withdraw()
    fTyp = [("","*")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo('○×プログラム','変換したい対象フォルダを今から選択してください')

    dir = tkinter.filedialog.askdirectory(initialdir = iDir)
    # path = dir + '\\'
    target_path = dir

    findFilesAndConvertToPdf(target_path, '*.png')
    findFilesAndConvertToPdf(target_path, '*.jpeg')
    findFilesAndConvertToPdf(target_path, '*.jpg')