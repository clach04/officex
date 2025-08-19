$PpFixedFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation
write-host $PpFixedFormat
$PowerPoint = New-Object -ComObject PowerPoint.application
# folderpath is a list of source directories, comma seperated
# Example:  $folderpath = "c:\fso\*", "c:\fso1\*"
# .pptx file(s) will go into the same directory as the source .ppt file
# By default, will use current directory as source and destination
# NOTE pptx files will be OVERWRITTEN
$folderpath = ".\"
$filetype ="*ppt"
Get-ChildItem -Path $folderpath -Include $filetype -recurse |
ForEach-Object `
{
    $path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))

    "Converting $path"
    "Using $PpFixedFormat"
    $Presentation = $PowerPoint.Presentations.open($_.fullname)

    $path += ".pptx"
    $Presentation.saveas($path, $PpFixedFormat)
    $Presentation.close()

    # move code
    #$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
    #
    #write-host $oldFolder
    #if(-not (test-path $oldFolder))
    #{
    #    new-item $oldFolder -type directory
    #}
    #
    #move-item $_.fullname $oldFolder

}
$PowerPoint.Quit()
$PowerPoint = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
