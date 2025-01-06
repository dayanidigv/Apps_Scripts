// Set the API key once
function setApiKey() {
  PropertiesService.getScriptProperties().setProperty('API_KEY', 'your_api_key_here');
  Logger.log('API Key stored successfully.');
}

// Get the API key
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (apiKey) {
    Logger.log('Retrieved API Key: ' + apiKey);
  } else {
    Logger.log('API Key not found.');
  }
}
