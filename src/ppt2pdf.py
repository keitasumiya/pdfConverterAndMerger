import os
import glob
import tkinter, tkinter.filedialog, tkinter.messagebox

def ppt2pdfIn (_target_path) :
    psCode = '''
    function ppt2pdf($ppt1, $ppt2){
        Write-Host "converting $($ppt2) files ..."

        $testExt = $scriptParentPath + "\*" + $ppt2
        if (Test-Path $testExt){
            $ppt = NEW-OBJECT -COMOBJECT POWERPOINT.APPLICATION
            $files = Get-ChildItem $scriptParentPath | Where-Object{$_.Name -match $ppt1}
            foreach($file in $files)
            {   
                try 
            {
                    # $doc = $ppt.Presentations.Open($file.FullName, [Type]::Missing, [Type]::Missing, [Type]::msoFalse)
                    $doc = $ppt.Presentations.Open($file.FullName, $False, $False, $False)
                    $doc.SaveAs([ref] $file.FullName.Replace($ppt2,".pdf"),[ref] 32)
                    $doc.Close()
                    Write-Host "$($file) was converted to a PDF successfully. "
                }
                catch
                {
                    Write-Host "[ERROR]$($file) was failed to covert to a PDF. "
                }
            }
            $ppt.Quit()

            $ppt = $null
            $doc = $null
            $file = $null
            $files = $null
        }
    }

    Write-Host "Start a conversion..."

    $scriptPath = $MyInvocation.MyCommand.Path

    $scriptParentPath = Split-Path -Parent $scriptPath

    ppt2pdf "ppt$" ".ppt"
    ppt2pdf "pptx$" ".pptx"
    ppt2pdf "pptm$" ".pptm"
    ppt2pdf "pps$" ".pps"
    ppt2pdf "ppsx$" ".ppsx"
    ppt2pdf "ppsm$" ".ppsm"

    # $doneFilePath = $scriptParentPath + "\_ppt2pdf_done.txt"
    # New-Item $doneFilePath -Force -Value done

    Write-Host "Finish a conversion..."
    '''

    fileBasename = r'_ppt2pdf.ps1'
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
    # target_path = r'C:\Users\ks\Documents\kisobutu\pdf-merger\code\ppt2pdf_make-ps\data'

    ppt2pdfIn(target_path)
