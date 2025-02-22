# Load the SCCM module
Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"

# Connect to the SCCM Site
cd "SCCM:A01"


Get-CMDevice
