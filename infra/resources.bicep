@description('Azure region for all resources')
param location string

@description('The azd environment name, used as a readable prefix in resource names')
param environmentName string

@description('Short unique token used to keep globally-unique resource names collision-free')
param resourceToken string

@description('Tags applied to every resource')
param tags object

var namePrefix = toLower(environmentName)
var nameSuffix = take(resourceToken, 5)

// Storage account name: lowercase alphanumeric only, max 24 chars, no hyphens
var storageAccountName = take('${replace(namePrefix, '-', '')}${nameSuffix}', 24)
var functionAppName = '${namePrefix}-onenote-proxy-${nameSuffix}'
var hostingPlanName = '${namePrefix}-onenote-proxy-plan-${nameSuffix}'
var appInsightsName = '${namePrefix}-onenote-proxy-appi-${nameSuffix}'

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: storageAccountName
  location: location
  tags: tags
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
}

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  tags: tags
  kind: 'web'
  properties: {
    Application_Type: 'web'
    IngestionMode: 'ApplicationInsights'
  }
}

resource hostingPlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: hostingPlanName
  location: location
  tags: tags
  kind: 'functionapp'
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  properties: {
    reserved: true
  }
}

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  tags: union(tags, { 'azd-service-name': 'api' })
  kind: 'functionapp,linux'
  properties: {
    serverFarmId: hostingPlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'Node|22'
      minTlsVersion: '1.2'
      ftpsState: 'Disabled'
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTSHARE'
          value: toLower(functionAppName)
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node'
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~22'
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsights.properties.ConnectionString
        }
        {
          name: 'PROXY_BASE_URL'
          value: 'https://${functionAppName}.azurewebsites.net'
        }
        // ENTRA_TENANT_ID / ENTRA_CLIENT_ID are filled in by
        // infra/hooks/postprovision.sh after this template creates the
        // Function App — Entra app registrations are Microsoft Graph
        // objects, not ARM/Bicep resources, so that step still shells out
        // to `az ad app create`.
      ]
    }
  }
}

output functionAppName string = functionApp.name
output functionAppHostName string = functionApp.properties.defaultHostName
output storageAccountName string = storageAccount.name
