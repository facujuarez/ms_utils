# General data
# Description: "Restore the last published major version of a document."

# Script execution steps
# 1. Connect to the source SharePoint Online Site
# 2. Get and Download the current version of the document
# 3. Download all versions of the document

# Clear and present
Clear-Host
Write-Host -f DarkCyan "============================================="
Write-Host -f DarkCyan "= Exportación de versiones de documentos ="
Write-Host -f DarkCyan "============================================="

Write-Host
Write-Host -f Magenta "Este proceso exporta todas las versiones previas de un documento en una biblioteca de documentos de SharePoint."
Write-Host

#Set Parameters
$SiteURL = "https://facujuarezdev.sharepoint.com/sites/GestionDocumentalDev"
$FileRelativeURL = "/sites/GestionDocumentalDev/Versions/Document.docx"
$DownloadPath = "C:\Temp"

try {

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
        Write-host -f Yellow "No se encontraron versiones."
    }
    
}
catch {
    Write-Host -f Red "Error al conectarse a SharePoint Online."
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host
Write-Host -f Blue "Ejecución del script finalizada."
Write-Host



