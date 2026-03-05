@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@secure()
param totalAgilityEndpoint string

@secure()
param totalAgilityApiKey string

@secure()
param totalAgilityAgentName string

@secure()
param totalAgilityAgentId string

@secure()
param totalAgilityTestUserName string

@secure()
param totalAgilityUseTestUser string

// Optional LLM configuration parameters.  Stored as strings so they
// can be passed from environment variables; defaults will be applied in
// the application code if they are empty or missing.
param totalAgilityTemperature string = ''
param totalAgilityUseSeed string = ''
param totalAgilitySeed string = ''

// Conversation history
param conversationHistoryMaxEntries string = ''

// Proactive notifications
@secure()
param notificationsBearerToken string = ''

param webAppSKU string

@maxLength(42)
param botDisplayName string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param identityName string = resourceBaseName
param storageAccountName string = toLower(replace('${resourceBaseName}stg', '-', ''))
param location string = resourceGroup().location

resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  location: location
  name: identityName
}

// Azure Storage account for persisting conversation references (Table Storage)
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: storageAccountName
  location: location
  kind: 'StorageV2'
  sku: {
    name: 'Standard_LRS'
  }
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
  }
}

// Enable Table service on the storage account
resource tableService 'Microsoft.Storage/storageAccounts/tableServices@2023-01-01' = {
  parent: storageAccount
  name: 'default'
}

// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
}

// Web App that hosts your bot
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure App Service from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~18' // Set NodeJS version to 18.x for your site
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'BOT_ID'
          value: identity.properties.clientId
        }
        {
          name: 'BOT_TENANT_ID'
          value: identity.properties.tenantId
        }
        { 
          name: 'BOT_TYPE'
          value: 'UserAssignedMsi' 
        }
        {
          name: 'TOTALAGILITY_ENDPOINT'
          value: totalAgilityEndpoint
        }
        {
          name: 'TOTALAGILITY_API_KEY'
          value: totalAgilityApiKey
        }
        {
          name: 'TOTALAGILITY_AGENT_NAME'
          value: totalAgilityAgentName
        }
        {
          name: 'TOTALAGILITY_AGENT_ID'
          value: totalAgilityAgentId
        }
        {
          name: 'TOTALAGILITY_TEST_USERNAME'
          value: totalAgilityTestUserName
        }
        {
          name: 'TOTALAGILITY_USE_TEST_USER'
          value: string(totalAgilityUseTestUser)
        }
        {
          name: 'TOTALAGILITY_TEMPERATURE'
          value: totalAgilityTemperature
        }
        {
          name: 'TOTALAGILITY_USE_SEED'
          value: totalAgilityUseSeed
        }
        {
          name: 'TOTALAGILITY_SEED'
          value: totalAgilitySeed
        }
        {
          name: 'CONVERSATION_HISTORY_MAX_ENTRIES'
          value: conversationHistoryMaxEntries
        }
        {
          name: 'NOTIFICATIONS_BEARER_TOKEN'
          value: notificationsBearerToken
        }
        {
          name: 'AZURE_STORAGE_CONNECTION_STRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=${environment().suffixes.storage}'
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    identityClientId: identity.properties.clientId
    identityResourceId: identity.id
    identityTenantId: identity.properties.tenantId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
output AZURE_STORAGE_ACCOUNT_NAME string = storageAccount.name
