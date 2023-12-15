# General data
# Username: ap-bp365@centralpuerto.com
# Password: (5_&{wGXypgG;qjW
# Backup site: https://centralpuerto.sharepoint.com/sites/DeprecatedSites
# Descripción: "Backup de los sitios de SharePoint Online depurados como parte de la implementación de las políticas de colaboración y gobierno de Central Puerto."

# Script execution steps
# 1. Ask for the environment to use (develop or production)
# 2. Connect to the source SharePoint Online Site
# 3. Get and Download the current version of the document
# 4. Get current document metadata
# 5. Download all versions of the document
# 6. Delete the document with all versions
# 7. Orderly upload document versions
# 8. Update the metadata to the document in its latest version

# Clear and present
Clear-Host
Write-Host "=============================================" -ForegroundColor Green
Write-Host "= Restauración de versiones de documentos de CAF =" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

Write-Host
Write-Host "Este proceso ..." -ForegroundColor Magenta
Write-Host

# Set main variables
$developTenant = "facujuarezdev"
$productionTenant = "cafpower"
$developSiteAlias = "GestionDocumentalDev"
$productionSiteAlias = "GestionDocumental"

# 1. Ask for the environment to use (develop, testing or production)
$environment = Read-Host "¿Desea usar el ambiente de producción o de desarrollo? (P/D)"
if ($environment -eq "P") {
  $tenantName = $productionTenant
  $siteAlias = $productionSiteAlias
  $siteUrl = "https://${tenantName}.sharepoint.com/sites/${siteAlias}"
}
elseif ($environment -eq "D") {
  $tenantName = $developTenant
  $siteAlias = $developSiteAlias
  $siteUrl = "https://${tenantName}.sharepoint.com/sites/${siteAlias}"
}
else {
  Write-Host "No se ha ingresado un valor válido. El script se cerrará." -ForegroundColor Red
  exit
}

# Ask for the document library to evaluate
# $documentLibraryName = Read-Host "Ingrese el nombre de la biblioteca de documentos a tratar (ej. 'Versiones')"
$documentLibraryName = "Versions"

# Ask for the document name to evaluate
# $documentName = Read-Host "Ingrese el nombre del documento a tratar (ej. 'Documento.docx')"
$documentName = "Document1.docx"

# Ask for the download path 
# $DownloadPath = Read-Host "Ingrese la ruta donde descargar los archivos (ej. 'C:\Temp')"
$DownloadPath = "C:\Temp"

Write-Host
Write-Host "Iniciando script..." -ForegroundColor Yellow
Write-Host

# Import PowerShell Modules
# Write-Host "Importando módulos de PnP PowerShell..." -ForegroundColor Yellow
# try {
#   Import-Module -Name "PnP.PowerShell" -ErrorAction Stop
# }
# catch {
#   Write-Host "Error al importar PnP PowerShell Module" -ForegroundColor Red
#   Write-Host $_.Exception.Message -ForegroundColor Red
# }

# 2. Connect to the source SharePoint Online Site
Write-Host "Conectando al sitio ${siteAlias}..." -ForegroundColor Yellow
try {
  Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
  
  # Get PnP Context to use in requests
  $Ctx = Get-PnPContext

  # Get PnP Connection to use in requests
  $sourceConnection = Get-PnPConnection

}
catch {
  Write-Host "Error al conectarse a SharePoint Online" -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
}

# 3. Get and Download the current version of the document
Write-Host "Descargando la versión actual del documento ${documentName}..." -ForegroundColor Yellow
try {
  # Get PnP Connection to use in request
  $sourceConnection = Get-PnPConnection

  #Get the File
  $FileRelativeURL = "/sites/${siteAlias}/${documentLibraryName}/${documentName}"
  $File = Get-PnPFile -Url $FileRelativeURL
  Write-Host $File

  # Get file as bytes
  $currentFileStream = (Get-PnPFile -Url "/sites/${siteAlias}/${documentLibraryName}/${documentName}" -AsFileObject -Connection $sourceConnection).OpenBinaryStream()
  Invoke-PnPQuery

  #Download File version to local disk
  # $CurrentVersionPathName = "$($DownloadPath)\$($Version.VersionLabel)_$($File.Name)"
  $CurrentVersionPathName = "$($DownloadPath)\$($documentName)"
  
  [System.IO.FileStream] $FileStream = [System.IO.File]::Open($CurrentVersionPathName, [System.IO.FileMode]::Create)
  $currentFileStream.Value.CopyTo($FileStream)
  $FileStream.Close()
  
  Write-Host -f Green "Version $($documentName) Downloaded to :" $CurrentVersionPathName 

}
catch {
  Write-Host "Error al descargar la versión actual del documento" -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
}

# 4. Get current document metadata
try {

  # Get File Versions
  $FileId = Get-PnPProperty -ClientObject $File -Property Id 
  Write-Host $FileId

}
catch {
  Write-Host "Error al " -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
}

# 5. Download all versions of the document
Write-Host "Descargando las versiones previas del documento ${documentName}..." -ForegroundColor Yellow
try {

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
      [System.IO.FileStream] $FileStream = [System.IO.File]::Open($VersionFileName, [System.IO.FileMode]::Create)
      $VersionStream.Value.CopyTo($FileStream)
      $FileStream.Close()
        
      Write-Host -f Green "Version $($Version.VersionLabel) descargada en :" $VersionFileName
    }
  }
  Else {
    Write-host -f Yellow "No se encontraron versiones previas."
  } 

}
catch {
  Write-Host "Error al descargar todas las versiones previas del documento." -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
}

# 6. Delete the document with all versions
Write-Host "Eliminando versión actual de la biblioteca de documentos ${documentLibraryName}..." -ForegroundColor Yellow
try {
  
  Remove-PnPFile -ServerRelativeUrl $FileRelativeURL

}
catch {
  Write-Host "Error al eliminar el documento de la biblioteca de documentos" -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
}


# try {
  

# }
# catch {
#   Write-Host "Error al " -ForegroundColor Red
#   Write-Host $_.Exception.Message -ForegroundColor Red
# }


# ------------------------------------------------------------------------------------------------------------
  
# Get the previous file version
# $previousVersion = Get-PnPFileVersion -Url "/sites/${siteAlias}/${documentLibraryName}/Document1.docx"
  
#$previousVersion = Get-PnPListItemVersion -List $documentLibraryName -Identity 10
#foreach ($previousVersionItem in $previousVersion) {
#  Write-Host "VersionLabel: $($previousVersionItem.VersionLabel)"
#}

# Download the previous major version content
# Get-PnPFile -Url "/sites/${siteAlias}/${documentLibraryName}/Document1.docx" -Path c:\temp -FileName Document1.docx -AsFile
# Write-Output "File downloaded to C:\Temp"
  
# # Delete the current file (including all minor versions)
# Remove-PnPFile -List $LibraryName -ItemId $ItemId
# Remove-PnPFile -ServerRelativeUrl "/sites/${siteAlias}/${documentLibraryName}/Document1.docx"
# Write-Output "File removed"

# Upload the previous major version content back to the library
# Add-PnPFile -List $LibraryName -ItemId $ItemId -Path ".\previousVersion.docx"
# Add-PnPFile -FileName Document11.docx -Folder "Versions" -Stream $stream -Values @{Modified = "12/14/2023" } 

# Write-Output "File reverted to previous major version (1.0) successfully!"

# $output = Get-PnPListItemVersion -List $documentLibraryName -Identity 16
# Get-PnPFile -List $documentLibraryName -ItemId 16 -VersionNumber 1 -Path ./previousVersion.docx

# Obtén información sobre el documento
# $file = Get-PnPFile -Url "/sites/${siteAlias}/${documentLibraryName}/Document1.docx"
# Revierte a la versión deseada (en este caso, la versión 4.0)
# Restore-PnPTenantRecycleBinItem -Identity $file.Versions[1.0].RecycleBinItemId -List "/sites/${siteAlias}/${documentLibraryName}"

# $output = Invoke-PnPSPRestMethod -Method Post -Url "/_api/web/GetFolderByServerRelativeUrl('/sites/${siteAlias}/${documentLibraryName}')/files('Document.docx')/unpublish(comment='Check-in comment for the unpublish operation.')"
# $output = Invoke-PnPSPRestMethod -Url "/_api/web/GetFolderByServerRelativeUrl('/sites/${siteAlias}/${documentLibraryName}')/files('Document.docx')/$value"
# $output = Invoke-PnPSPRestMethod -Method Post -Url "/_api/web/GetFileByServerRelativeUrl('/sites/${siteAlias}/${documentLibraryName}/Document.docx')/unpublish()"
  
# $output = Invoke-PnPSPRestMethod -Url "/_api/web/GetFileByServerRelativePath(decodedurl='/sites/${siteAlias}/${documentLibraryName}/Document1.docx')/versions()" -Content $item
# Write-Host "Previous version URL: " $previousVersion.Url
# Write-Host "Previous version Label: "$previousVersion.VersionLabel
# Write-Host "Previous version Stream: " $previousVersionBinaryStream
# Write-Host "Current version : "$currentVersion

# Retrieves the version history of a file, not including its current version
# $fileVersion = Get-PnPFileVersion -Url "/sites/${siteAlias}/${documentLibraryName}/3.Documento%20de%20prueba.docx" 
# Write-Host "Datos de versiones 3.Documento%20de%20prueba.docx..." -ForegroundColor Yellow
# $fileVersion
# Write-Host