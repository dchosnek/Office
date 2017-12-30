<#
.SYNOPSIS
  Unlock protected sheets in an Excel file.
.DESCRIPTION
  Unlocks protected sheets, even those protected by a password.
.INPUTS
  A list of filenames to unlock.
.OUTPUTS
  For each input file, an unlocked output file is created.
.NOTES
  Version:        1.0
  Author:         Doron Chosnek
  Creation Date:  March 2017
  Purpose/Change: Initial script development
#>


# The general process here, which is documented many places on the web, is:
# 1) Copy the original file with .xlsx or .xlsm extension to .zip extension
# 2) Open each file in the xl/worksheets directory of the .zip file
# 3) Search for an XML tag 'sheetProtection' and remove it
# 4) Rename the zip file back to the original extension


# pass one or more filenames to the script
param( 
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][string[]]$Filename 
)


# Thanks to http://techibee.com/powershell/reading-zip-file-contents-without-extraction-using-powershell/2152


# you must have DotNet vesion 4.5 at minimum for the zip functions in here to work
$dotnetversion = [Environment]::Version            
if(!($dotnetversion.Major -ge 4 -and $dotnetversion.Build -ge 30319)) {
    Write-Error "Microsoft DotNet Framework 4.5 required."
    exit(1)            
}

# https://social.technet.microsoft.com/Forums/windowsserver/en-US/4512d72a-7dec-4b27-8a83-db31f84334fb/overwrite-zip-contents

[void] [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
[void] [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression')


foreach($locked_name in $Filename)
{
    # copy the original file to zip file extension
    $locked_file = Get-ChildItem $locked_name
    $zip_name = $locked_file.FullName -replace '\.xls\w*$', '_unlocked.zip'
    Copy-Item $locked_file -Destination $zip_name

    # create a temporary directory
    $dir_name = $zip_name -replace '\.zip$', ''
    New-Item $dir_name -ItemType Directory | Out-Null

    # extract only the files from the zip that we care about
    $Archive = [System.IO.Compression.ZipFile]::Open($zip_name, [System.IO.Compression.ZipArchiveMode]::Update)
    $xml_files = $Archive.Entries | ? {$_.FullName -match 'xl/worksheets/.+xml$'} 
    foreach($xf in $xml_files) 
    {
        $name_in_dir = "$($dir_name)\$($xf.Name)"
        $name_in_zip = $xf.FullName

        # extract to temporary directory
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($xf, $name_in_dir, $true)

        # change file if necessary; if the file is changed, dump it back into the zip file
        $contents = [string](Get-Content $name_in_dir)
        if ( $contents -match '<sheetProtection.+?/>')
        {
            $new_contents = $contents -replace '<sheetProtection.+?/>', ''
            $new_contents | Set-Content $name_in_dir
            $entry = $Archive.GetEntry($name_in_zip)<##>
            if($entry) { $entry.Delete() }
            [void] [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $name_in_dir, $name_in_zip)
            Write-Host "Unlocked", ($xf.Name -replace '.xml', '') -ForegroundColor Green
        }
    }    

    # close the zip that was opened in update mode
    $Archive.Dispose()

    # remove the temporary directory that was created
    Remove-Item $dir_name -Recurse -Force

    # rename the zip file back to the original extension
    $unlocked_name = $dir_name + $locked_file.Extension
    Move-Item $zip_name -Destination $unlocked_name

}
