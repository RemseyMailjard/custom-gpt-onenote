targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the azd environment. Used as the deployment identity (one per tenant/customer) and to derive resource names.')
param environmentName string

@minLength(1)
@description('Azure region for all resources')
param location string

var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = {
  'azd-env-name': environmentName
}

resource rg 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: 'rg-${environmentName}'
  location: location
  tags: tags
}

module resources 'resources.bicep' = {
  name: 'resources'
  scope: rg
  params: {
    location: location
    environmentName: environmentName
    resourceToken: resourceToken
    tags: tags
  }
}

output AZURE_LOCATION string = location
output RESOURCE_GROUP_NAME string = rg.name
output FUNCTION_APP_NAME string = resources.outputs.functionAppName
output FUNCTION_APP_HOSTNAME string = resources.outputs.functionAppHostName
output PROXY_BASE_URL string = 'https://${resources.outputs.functionAppHostName}'
output AZURE_STORAGE_ACCOUNT_NAME string = resources.outputs.storageAccountName
