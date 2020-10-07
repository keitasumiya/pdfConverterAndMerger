import os
import glob
import tkinter, tkinter.filedialog, tkinter.messagebox

def xls2pdfIn (_target_path) :
    psCode = '''
    function xls2pdf($xls1, $xls2){
        Write-Host "converting $($xls2) files ..."

        $testExt = $scriptParentPath + "\*" + $xls2
        if (Test-Path $testExt){
            $xls = NEW-OBJECT -COMOBJECT Excel.APPLICATION
            $xls.Visible = $False
            $files = Get-ChildItem $scriptParentPath | Where-Object{$_.Name -match $xls1}
            foreach($file in $files)
            {   
                try 
            {
                    $doc = $xls.Workbooks.Open($file.FullName, $False, $True)
                    $doc.ExportAsFixedFormat($Type, $file.FullName.Replace($xls2,".pdf"), $Quality, $IncludeDocProperties, $IgnorePrintAreas)
                    $doc.Close()
                    Write-Host "$($file) was converted to a PDF successfully. "
                }
                catch
                {
                    Write-Host "[ERROR]$($file) was failed to covert to a PDF. "
                }
            }
            $xls.Quit()

            $xls = $null
            $doc = $null
            $file = $null
            $files = $null
        }
    }


    Write-Host "Start a conversion..."

    $XlFixedFormatType = "Microsoft.Office.Interop.Excel.XlFixedFormatType" -as [type]
    $XlFixedFormatQuality = "Microsoft.Office.Interop.Excel.XlFixedFormatQuality" -as [type]

    $Type = $XlFixedFormatType::xlTypePDF
    $Quality = $XlFixedFormatQuality::xlQualityStandard
    $IncludeDocProperties = $True
    $IgnorePrintAreas = $False

    $scriptPath = $MyInvocation.MyCommand.Path

    $scriptParentPath = Split-Path -Parent $scriptPath

    xls2pdf "xls$" ".xls"
    xls2pdf "xlsx$" ".xlsx"
    xls2pdf "xlsm$" ".xlsm"

    # $doneFilePath = $scriptParentPath + "\_xls2pdf_done.txt"
    # New-Item $doneFilePath -Force -Value done

    Write-Host "Finish a conversion..."
    '''

    fileBasename = r'_xls2pdf.ps1'
    path_w = _target_path + '\\' + fileBasename
    with open(path_w, mode='w') as f:
        f.write(psCode)

    # os.system(r'powershell -Command .\data\_xls-xlsx-xlsm2pdf.ps1')
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

    # target_path = r'.\data'
    # target_path = r'C:\Users\ks\Documents\kisobutu\pdf-merger\code\xls2pdf_make-ps\data'
    target_path = dir
    xls2pdfIn(target_path)
