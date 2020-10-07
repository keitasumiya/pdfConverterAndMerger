import os
import glob
import tkinter, tkinter.filedialog, tkinter.messagebox

def doc2pdfIn (_target_path) :
    psCode = '''
    function doc2pdf($doc1, $doc2) {
        Write-Host "converting $($doc2) files ..."

        $testExt = $scriptParentPath + "\*" + $doc2
        if (Test-Path $testExt){
            $word = NEW-OBJECT -COMOBJECT WORD.APPLICATION

            $files = Get-ChildItem $scriptParentPath | Where-Object{$_.Name -match $doc1}
            foreach($file in $files)
            {   
                try 
            {
                    $doc = $word.Documents.OpenNoRepairDialog($file.FullName)
                    $doc.SaveAs([ref] $file.FullName.Replace($doc2,".pdf"),[ref] 17)
                    $doc.Close()
                    Write-Host "$($file) was converted to a PDF successfully. "
                }
                catch
                {
                    Write-Host "[ERROR]$($file) was failed to covert to a PDF. "
                }
            }
            $word.Quit()

            $word = $null
            $doc = $null
            $file = $null
            $files = $null
        }
    }


    Write-Host "Start a conversion..."

    $scriptPath = $MyInvocation.MyCommand.Path

    $scriptParentPath = Split-Path -Parent $scriptPath

    doc2pdf "doc$" ".doc"
    doc2pdf "docx$" ".docx"
    doc2pdf "docm$" ".docm"

    # $doneFilePath = $scriptParentPath + "\_doc2pdf_done.txt"
    # New-Item $doneFilePath -Force -Value done

    Write-Host "Finish a conversion..."
    '''

    
    fileBasename = r'_doc2pdf.ps1'
    path_w = _target_path + '\\' + fileBasename

    with open(path_w, mode='w') as f:
        f.write(psCode)

    # os.system(r'powershell -Command .\data' + '\\' + fileBasename)
    # os.system(r'powershell -Command ' + _target_path + '\\' + fileBasename)
    os.system(r'powershell -NoProfile -ExecutionPolicy Unrestricted ' + _target_path + '\\' + fileBasename)

    os.remove(path_w)

if __name__ == '__main__':
    # フォルダ選択ダイアログの表示
    root = tkinter.Tk()
    root.withdraw()
    fTyp = [("","*")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo('○×プログラム','変換したい対象フォルダを今から選択してください')

    dir = tkinter.filedialog.askdirectory(initialdir = iDir)
    target_path = dir

    # target_path = r'.\data'
    # target_path = r'C:\Users\ks\Documents\kisobutu\pdf-merger\code\doc2pdf_make-ps\data'

    doc2pdfIn(target_path)
