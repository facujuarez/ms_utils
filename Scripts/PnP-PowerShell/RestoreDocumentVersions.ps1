# General data
# Description: "Restore the last published major version of a document."

# Script execution steps
# 1. Ask for the environment to use (develop or production)
# 2. Connect to the source SharePoint Online Site
# 3. Get and Download the current version of the document
# 4. Get current document metadata
# 5. Download all versions of the document
# 6. Delete the document with all versions
# 7. Orderly upload document versions and update the metadata to the latest version

# Clear and present
Clear-Host
Write-Host -f DarkCyan "============================================="
Write-Host -f DarkCyan "= Restauración de versiones de documentos ="
Write-Host -f DarkCyan "============================================="

Write-Host
Write-Host -f Magenta "Este proceso restaura la última versión MAYOR de un documento en una biblioteca de documentos de SharePoint."
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
  Write-Host -f Red "No se ha ingresado un valor válido. El script se cerrará."
  exit
}

# Ask for the document library to evaluate
$documentLibraryName = Read-Host "Ingrese el nombre de la biblioteca de documentos a tratar (ej. 'Versiones')"

# Ask for the server relative URL of the document to evaluate
$fileSiteRelativeURL = Read-Host "Ingrese la ruta relativa del documento a tratar (ej. '${documentLibraryName}/Documento.docx')"

# Ask for the document library to evaluate
$targetVersion = Read-Host "Ingrese el número de la versión major objetivo a restaurar (ej. '1.0')"

# Ask for the download path 
# $DownloadPath = Read-Host "Ingrese la ruta donde descargar los archivos (ej. 'C:\Temp')"
$DownloadPath = "C:\Temp"

Write-Host
Write-Host -f Blue "Iniciando script..."
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
Write-Host -f Yellow "Conectando al sitio ${siteAlias}..."
try {
  Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
  
  # Get PnP Context to use in requests
  $ctx = Get-PnPContext

  # Get PnP Connection to use in requests
  $sourceConnection = Get-PnPConnection

}
catch {
  Write-Host -f Red "Error al conectarse a SharePoint Online."
  Write-Host $_.Exception.Message -ForegroundColor Red
}

# 3. Get and Download the current version of the document
Write-Host
Write-Host -f Yellow "Descargando la versión actual del documento..."
try {
  # Get PnP File
  $file = Get-PnPFile -Url $fileSiteRelativeURL -AsFileObject -Connection $sourceConnection
  $documentName = $file.Name

  # Get file as bytes
  $currentFileStream = $file.OpenBinaryStream()
  Invoke-PnPQuery

  #Download File version to local disk
  $CurrentVersionPathName = "$($DownloadPath)\$("Original")_$($documentName)"
  
  [System.IO.FileStream] $fileStream = [System.IO.File]::Open($CurrentVersionPathName, [System.IO.FileMode]::Create)
  $currentFileStream.Value.CopyTo($fileStream)
  $fileStream.Close()
  
  Write-Host -f Green "Version $($documentName) descargada como: " $CurrentVersionPathName 

}
catch {
  Write-Host -f Red "Error al descargar la versión actual del documento."
  Write-Host -f Red $_.Exception.Message
}

# 4. Get current document metadata
Write-Host
Write-Host -f Yellow "Obteniendo la metadata del documento ${documentName}..."
try {

  # Get PnP File as list item
  $file = Get-PnPFile -Url $fileSiteRelativeURL -AsListItem

  $fileProperties = $file.FieldValues
  Write-Host -f Yellow "Cantidad de propiedades encontradas: " $fileProperties.Count
  
  $filePropertiesKeysValues = @{}

  # Set list item properties to upgrade
  foreach ($property in $fileProperties.Keys) {

    switch ($property) {
      # "Aprobador" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupValue)} }
      # "AprobadorTexto" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property]) } }
      # "Autor" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupValue)} }
      # "AutorTexto" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "CargoAprobador" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "CargoAutor" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "CargoRevisorPrincipal" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "CategoriaLookup" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupId)} }
      # "CodigoCalidad" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "Componente" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      "Created" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "Descripcion" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "EstadoLookup" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupId)} }
      # "FechaAprobador" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property]) } }
      # "FechaAutor" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "FechaFirma" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "FechaDocumento" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "FechaRevisorPrincipal" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "DocumentoFirmado" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "IsFileLocked" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "Modified" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} } #Personalizar al de la ultima version major
      # "ModuloLookup" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupId)} }
      # "ProyClienteArea" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "RegistroRevision" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "RequiereRR" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property]) } }
      # "RevisionAdmon" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "RevisorAdministracion" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupValue)} }
      # "RevisorAdministracionTexto" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "RevisorPrincipal" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupValue)} }
      # "RevisorPrincipalTexto" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "Revisores" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "SistemaTematicaProd" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "ClaseLookup" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupId)} }
      # "TipoDocumentoLookup" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupId)} }
      # "TipoPlantillaLookup" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupId)} }
      # "Title" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "TituloDocumento" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "VersionTexto" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "WordApprovedToProducing" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "WordMoficicacion" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property])} }
      # "Author" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupValue)} }
      # "Editor" { $filePropertiesKeysValues += @{${property} = $($fileProperties[$property].LookupValue)} }
      
      Default {}
    }
  }

  # Print list item properties to upgrade
  Write-Host $filePropertiesKeysValues

}
catch {
  Write-Host -f Red "Error al obtener la metadata del documento."
  Write-Host -f Red $_.Exception.Message
}

# 5. Download all versions of the document
Write-Host
Write-Host -f Yellow "Descargando las versiones previas del documento ${documentName}..."
try {

  # Get PnP File
  $file = Get-PnPFile -Url $fileSiteRelativeURL
  
  #Get File Versions
  $fileVersions = Get-PnPProperty -ClientObject $file -Property Versions
  Write-Host -f Yellow "Cantidad de versiones encontradas: " $fileVersions.Count

  If ($fileVersions.Count -gt 0) {
    Foreach ($version in $fileVersions) {

      #Frame File Name for the Version
      $versionFileName = "$($DownloadPath)\$($version.VersionLabel)_$($file.Name)"

      #Get Contents of the File Version
      $versionStream = $version.OpenBinaryStream()
      $ctx.ExecuteQuery()
  
      #Download File version to local disk
      [System.IO.FileStream] $fileStream = [System.IO.File]::Open($versionFileName, [System.IO.FileMode]::Create)
      $versionStream.Value.CopyTo($fileStream)
      $fileStream.Close()
        
      Write-Host -f Green "Version $($version.VersionLabel) descargada como: " $versionFileName
    }
  }
  Else {
    Write-host -f Yellow "No se encontraron versiones previas."
  } 

}
catch {
  Write-Host -f Red "Error al descargar todas las versiones previas del documento."
  Write-Host -f Red $_.Exception.Message
}

# 6. Delete the document with all versions
Write-Host
Write-Host "Eliminando versión actual y previas del documento ${documentName}..." -ForegroundColor Yellow
try {
  
  # Delete file from Document library
  Remove-PnPFile -SiteRelativeUrl "$($documentLibraryName)/$($documentName)" -Connection $sourceConnection -Force
  Write-Host -f Green "Documento $($documentLibraryName)/$($documentName) removido."

}
catch {
  Write-Host -f Red "Error al eliminar el documento de la biblioteca de documentos."
  Write-Host -f Red $_.Exception.Message
}

Write-Host
Write-Host -f Magenta "Preparando la restauración de versiones previas..."
Start-Sleep -Seconds 2

# 7. Orderly upload document versions
Write-Host
Write-Host "Restaurando versiones previas del documento ${documentName} en la biblioteca ${documentLibraryName} ..." -ForegroundColor Yellow
try {
  
  If ($fileVersions.Count -gt 0) {
    Foreach ($version in $fileVersions) {

      # Frame File Name for the Version
      $versionFilePath = "$($DownloadPath)\$($version.VersionLabel)_$($documentName)"
      $documentFilePath = "$($DownloadPath)\$($documentName)"

      # Rename version document to original name
      Rename-Item -Path $versionFilePath -NewName $documentName

      # Valida si es la major version final
      If ($version.VersionLabel.Contains($targetVersion)) {

        # Add file version to document library
        Add-PnPFile -Path $documentFilePath -Folder $documentLibraryName -Publish -PublishComment $version.CheckInComment -ContentType "Documento de calidad" 
        Write-Host -f Green "Versión objetivo $($version.VersionLabel) restaurada."

        Write-Host
        Write-Host -f Blue "Terminando proceso de restauración..."
        break
      }
      Else {

        # Add file version to document library
        Add-PnPFile -Path $documentFilePath -Folder $documentLibraryName -CheckInComment $version.CheckInComment -ContentType "Documento de calidad"
        Write-Host -f Green "Version $($version.VersionLabel) restaurada."

        # Add properties to file as list item
        $fileListItem = Get-PnPFile -Url $fileSiteRelativeURL -AsListItem
        Set-PnPListItem -List $documentLibraryName -Identity $fileListItem.Id -Values $filePropertiesKeysValues -UpdateType UpdateOverwriteVersion -Force
        Write-Host -f Green "Propiedades actualizadas para version $($version.VersionLabel)"
      }

      # Revert rename version document to original name
      Rename-Item -Path $documentFilePath -NewName "$($version.VersionLabel)_$($documentName)"

    }
  }
  Else {
    Write-host -f Yellow "No se encontraron versiones previas."
  } 

}
catch {
  Write-Host -f Red "Error al restaurar versiones previas del documento."
  Write-Host -f Red $_.Exception.Message
}

Write-Host
Write-Host -f Blue "Ejecución del script finalizada."
Write-Host

# Write-Host "..." -ForegroundColor Yellow
# try {
  

# }
# catch {
#   Write-Host "Error al " -ForegroundColor Red
#   Write-Host $_.Exception.Message -ForegroundColor Red
# }
