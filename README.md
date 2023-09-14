### Cert
You must upload the .pfx cert file to the Function App sert settings page

### Configuration Settings
The following configuration settings are required in the Function Configuration settings. 
```json
[  {
    "name": "TenantId",
    "value": "",
    "slotSetting": false
  },
  {
    "name": "ClientId",
    "value": "",
    "slotSetting": false
  },
  {
    "name": "SiteUrl",
    "value": "",
    "slotSetting": false
  },
    {
    "name": "Thumbprint",
    "value": "",
    "slotSetting": false
  },
  {
    "name": "StorageConnectionString",
    "value": "",
    "slotSetting": false
  },
  {
    "name": "StorageContainerName",
    "value": "",
    "slotSetting": false
  },
  {
    "name": "WEBSITE_LOAD_CERTIFICATES",
    "value": "*",
    "slotSetting": false
  }
]
```