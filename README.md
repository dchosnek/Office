# Microsoft Office Scripts

This is a collection of scripts that simplify some tasks with different Microsoft Office programs.

## Get-PptLinks.ps1

Users can paste ***linked*** data (like cells from Excel) into PowerPoint so that the presentation always contains the latest data from a spreadsheet. The issue is that it's difficult to view a summary of all linked data pasted into a presentation. This script will display the links and slide numbers for each in a window with sortable columns.

### How to run

The script only has one argument: the PowerPoint filename.
```
Get-PptLinks.ps1 -File sample.pptx
```
The filename can be passed by the pipeline
```
'sample.pptx' | Get-PptLinks.ps1
```
Or analyze every file in a directory:
```
Get-ChildItem *.pptx | % { Get-PptLinks.ps1 $_ }
```

## Unlock-ExcelSheet.ps1

Individual sheets in an Excel file can be locked and password protected. This script defeats the password protection of every worksheet in the specified Excel workbook and creates a new copy of the workbook with all sheets unlocked.

The script uses a process widely documented on the internet that takes advantage of the fact that Microsoft Office files are essentially zip archives comprised of XML files. The process is:

The general process here, which is documented many places on the web, is:
1) Copy the original file with .xlsx or .xlsm extension to .zip extension
2) Open each file in the xl/worksheets directory of the .zip file
3) Search for an XML tag 'sheetProtection' and remove it
4) Rename the zip file back to the original extension

This script does not work if the workbook itself is password-protected.