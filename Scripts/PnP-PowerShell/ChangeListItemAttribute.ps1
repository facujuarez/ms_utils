# General data
# Description: "Restore the last published major version of a document."

# Script execution steps
# 1. Connect to the source SharePoint Online Site
# 2. Set values to update
# 3. Update list item

# Clear and present
Clear-Host
Write-Host -f DarkCyan "============================================="
Write-Host -f DarkCyan "= Actualización de atributos de elementos de lista ="
Write-Host -f DarkCyan "============================================="

Write-Host
Write-Host -f Magenta "Este proceso actualiza valores de atributos de lista de SharePoint."
Write-Host

$developTenant = "facujuarezdev"
$developSiteAlias = "GestionDocumentalDev"

$tenantName = $developTenant
$siteAlias = $developSiteAlias
$siteUrl = "https://${tenantName}.sharepoint.com/sites/${siteAlias}"

Write-Host
Write-Host "Iniciando script..." -ForegroundColor Magenta
Write-Host

Write-Host "Conectando al sitio ${siteAlias}..." -ForegroundColor Yellow
try {
    # Conecta a SharePoint
    Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
}
catch {
    Write-Host "Error al conectarse a SharePoint Online" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host "Actualizando propiedades del documento..." -ForegroundColor Yellow
try {
    # Obtiene el elemento de lista
    $item = Get-PnPListItem -List "Test" -ID 1
    Write-Host $item.FieldValues

    # Establece la nueva fecha de creación
    $newCreatedDate = (Get-Date -Year 2023 -Month 07 -Day 20)
    Write-Host $newCreatedDate

    # Actualiza el elemento de lista
    Set-PnPListItem -List "Test" -Identity 1 -Values @{"Created" = $newCreatedDate; }

    Write-Host "Propiedades actualizadas correctamente." -ForegroundColor Green
}
catch {
    Write-Host "Error al conectarse a SharePoint Online" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host
Write-Host "Fin ejecución del script." -ForegroundColor Magenta
Write-Host



