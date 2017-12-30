<#
.SYNOPSIS
Displays the links for linked objects in a given PPT file (if there are any).

.DESCRIPTION
This script will display the short and long link for any linked objects on
every PPT slide in a PPT deck. Links are displayed in a new windows that 
allows sorting and filtering, but the script can be easily modified to display
the links in a different manner.

.PARAMETER Filename
This required parameter specifies the file to be examined and can be passed
through the pipeline.

.NOTES
  Version:        1.0
  Author:         Doron Chosnek
  Creation Date:  December 2017
  Purpose/Change: Initial script development

.EXAMPLE
 Get-PptLinks.ps1 -Filename sample.pptx
    Displays links for all linked objects in the file sample.pptx.
.EXAMPLE
 sample.pptx | Get-PptLinks.ps1
    Displays links for all linked objects in the file sample.pptx.
#>


param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)][string]$File
)

# ==========================================
# open the PPT file
# ------------------------------------------

$filename = (Resolve-Path $File).Path
$powerpoint = New-Object -ComObject powerpoint.application
#$Application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$powerpoint.Visible = 1
$preso = $powerpoint.Presentations.Open($filename)

# ==========================================
# examine each link in each slide
# ------------------------------------------

$summary = @()

foreach($currentslide in $preso.Slides) {
    foreach($currentshape in $currentslide.Shapes) {
		$progress = ($currentslide.SlideNumber / $preso.Slides.Count * 100)
        Write-Progress -Id 1 -Activity "Analyzing shapes on slide $($currentslide.SlideNumber) of $($preso.Slides.Count)" `
                       -Status "$($currentshape.Name)" -PercentComplete $progress -CurrentOperation "$progress% complete"
        if($currentshape.LinkFormat.SourceFullName)
        {
            $a = "" | select Slide, Name, ShortLink, LongLink
            $a.Slide = $currentslide.SlideNumber
            $a.Name = $currentshape.Name
            $a.LongLink = $currentshape.LinkFormat.SourceFullName
            $a.ShortLink = $a.LongLink -replace '.*\\', ""
            $summary += $a

        }
    }
}

if($summary.Count -gt 0) {
    $summary | Out-GridView -Title $filename
}

# ==========================================
# close PowerPoint
# ------------------------------------------

# This seems to be the safest way to close PowerPoint. It won't close
# any windows that have unsaved changes in them. It doesn't close ALL 
# PowerPoint windows; it only closes the most recently opened window 
# (the one opened by this script). If you run Get-Process you'll find
# only one POWERPNT process even if you have multiple instances open.
# So CloseMainWindow() will only close the active  window. This is
# different than Notepad, which opens a new process for each window.

Get-Process POWERPNT | % { $_.CloseMainWindow() | Out-Null }

# ==========================================
# Show a list of files linked to this PPT
# ------------------------------------------

$files = @()
$summary | select -ExpandProperty ShortLink | % {
    if($_ -match '(.*?)!')
    {
        $files += $Matches[1]
    }
}

if($files.Count -gt 0) {
    Write-Host "This PPT is linked to the following files:" -ForegroundColor Green
    $files | Select-Object -Unique
}