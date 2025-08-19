# OfficeX Conversion Scripts

Convert old Microsoft Office documents to new
Use Microsoft Office to convert old office formats into (zip) X format.
I.e. Batch doc to docx, ppt to pptx, xls to xlsx conversion scripts.

This scripts expects:

  * Microsoft Windows
  * Microsoft Office that can read old documents (note possible restrictions with files pre 1998, see LibreOffice as potential backup)
  * PowerShell

There are alternatives see later.

## Setup

Read through https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies before proceeding!

 1. Open PowerShell
 2. Sanity chech current policy:

        Get-ExecutionPolicy -List

 3. Allow scripts to be ran:

        Set-ExecutionPolicy -ExecutionPolicy  Unrestricted  -Scope  Process

    For this shell/session ONLY!


## Conversion PowerShell Scripts

### Excel

TODO Consider adding as first line top of script, `Add-Type -AssemblyName Microsoft.Office.Interop.Excel`
TODO Look at using `$excel.visible = $false`

`.\ConvertXLS.ps1`

Original from https://gist.github.com/riskeez/096f3ee6bc23d35ed7730bbd36b33c44
Also see original-original https://gist.github.com/gabceb/954418

Assuming setup complete, from PowerShell window issue:

    .\SavePowerpointPptAsPptx.ps1

Assuming .xls files to convert are in the current directory.

WARNING existing xlsx files will cause Excel to pause and prompt (NOTE dialog box might be hidden behind PowerShell Window).

### Powerpoint

TODO Consider adding as first line top of script, `Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint`

`.\SavePowerpointPptAsPptx.ps1`

Original from https://dlairman.wordpress.com/2013/01/15/convert-ppt-to-pptx-using-powershell/

Assuming setup complete, from PowerShell window issue:

    .\SavePowerpointPptAsPptx.ps1

Assuming .ppt files to convert are in the current directory.

WARNING pptx files will be OVERWRITTEN!


### Word

`.\SaveWordDocAsDocx.ps1`

Original from https://web.archive.org/web/20150508085022/http://blogs.technet.com/b/heyscriptingguy/archive/2010/06/22/hey-scripting-guy-how-can-i-use-windows-powershell-2-0-to-convert-doc-files-to-docx-files.aspx


Assuming setup complete, from PowerShell window issue:

    .\SaveWordDocAsDocx.ps1

Assuming .doc files to convert are in the current directory.

WARNING docx files will be OVERWRITTEN!


## Alternatives

### Alternatives - Python

Python implementations also use COM.

```python
import win32com

w = win32com.gencache.EnsureDispatch("Word.Application")
doc = w.Documents.Open(in_filename)  # ???????doc
doc.SaveAs2(in_filename + "x", 12)  # I have no idea what the 12 literal constant is for...
#doc.SaveAs(in_filename + "x"_path, 12, False, "", True, "", False, False, False, False)
doc.Close()
w.Quit()

```

  * https://github.com/JPomichael/doc2docx/blob/master/doc2docx.py
  * https://github.com/luibkin/doc2docx/blob/master/doc2docx.py
  * https://github.com/baifengbai/doc2docx/blob/master/bl/docx.py
  * https://github.com/ryo-sasa/doc2docx/blob/main/doc2docx.py
  * https://github.com/rooooolin/doc2docx
  * https://github.com/zhouyunfan22/doc2docx/blob/master/main.py
  * https://github.com/liruishaer/doc2docx/blob/master/win32_doc2docx.py
  * https://github.com/ShuBo6/py-doc2docx/blob/main/main.py
  * https://github.com/yuzijiano/read_docx-doc2docx/blob/master/%E5%B0%86%E6%89%80%E6%9C%89%E7%9A%84doc%E5%A4%8D%E5%88%B6%E4%B8%80%E4%BB%BD%E5%8F%98%E6%88%90docx.py
  * https://github.com/hcxss/docx_batch_handle/blob/main/doc2docx.py

### Alternatives - LibreOffice / OpenOffice

Works cross platform, may not preserve formatting/display exactly.

    'soffice' + ' --headless --convert-to docx ' + in_filename + ' --outdir ' + out_path

  * https://github.com/Done-1026/doc2docx/blob/master/divide_doc.py
