#Set Parameters
$SiteURL = "https://facujuarezdev.sharepoint.com/sites/GestionDocumentalDev"
$FileRelativeURL = "/sites/GestionDocumentalDev/Versions/Document1.docx"
$DownloadPath = "C:\Temp"
 
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
$Ctx = Get-PnPContext
 
#Get the File
$File = Get-PnPFile -Url $FileRelativeURL

#Get File Versions
$FileVersions = Get-PnPProperty -ClientObject $File -Property Versions

If ($FileVersions.Count -gt 0) {
    Foreach ($Version in $FileVersions) {
        #Frame File Name for the Version
        $VersionFileName = "$($DownloadPath)\$($Version.VersionLabel)_$($File.Name)"

        #Get Contents of the File Version
        $VersionStream = $Version.OpenBinaryStream()
        $Ctx.ExecuteQuery()
  
        #Download File version to local disk
        [System.IO.FileStream] $FileStream = [System.IO.File]::Open($CurrentVersionFileName, [System.IO.FileMode]::Create)
        $VersionStream.Value.CopyTo($FileStream)
        $FileStream.Close()
        
        Write-Host -f Green "Version $($Version.VersionLabel) Downloaded to :" $VersionFileName        

    }
}
Else {
    Write-host -f Yellow "No Versions Found!"
}

# Delete the current file (including all minor versions)
# Remove-PnPFile -List $LibraryName -ItemId $ItemId
# Remove-PnPFile -ServerRelativeUrl $FileRelativeURL
# Write-Output "File removed"



