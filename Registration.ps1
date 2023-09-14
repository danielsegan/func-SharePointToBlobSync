# -== PnP ==-
# SDK - https://pnp.github.io/pnpcore/using-the-sdk/readme.html
# Function App Sample - https://github.com/pnp/pnpcore/tree/dev/samples/Demo.AzureFunction.OutOfProcess.AppOnly
# PWSH - https://pnp.github.io/powershell/index.html

$SITE_URL="https://seganassociates.sharepoint.com/sites/TestSite"

# Register PnP Auth w/ SP Online Tennant 
Register-PnPManagementShellAccess

# Connect 
Connect-PnPOnline seganassociates.sharepoint.com 

# Register App / This will output certs
$result = Register-PnPAzureADApp -ApplicationName "ClipWizzardApp3" -Tenant seganassociates.onmicrosoft.com -OutPath . -DeviceLogin -GraphApplicationPermissions "Sites.FullControl.All" -SharePointApplicationPermissions "Sites.FullControl.All" -CertificatePassword (ConvertTo-SecureString -String "password" -AsPlainText -Force)

# Test App Reg Accees
Connect-PnPOnline -Url $SITE_URL -ClientId 8f356b98-7652-4703-a064-6073e2e6a102 -Tenant seganassociates.onmicrosoft.com -CertificatePath ClipWizzardApp.pfx

