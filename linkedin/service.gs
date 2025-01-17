const CLIENT_ID = 'your_linkedin_client_id';
const CLIENT_SECRET = 'your_linkedin_client_secret';

/**
 * Authorizes and makes a request to the LinkedIn API.
 */
function run() {
  var service = getService_();
  if (service.hasAccess()) {
    var url = 'https://api.linkedin.com/v1/people/~?format=json';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
      authorizationUrl);
  }
}

function runMe() {
  var service = getService_();
  Logger.log(service.getAccessToken());
  if (service.hasAccess()) {
    var url = 'https://api.linkedin.com/v2/userinfo';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
      authorizationUrl);
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService_().reset();
}

/**
 * Configures the service.
 */
function getService_() {
  return OAuth2.createService('LinkedIn')
    // Set the endpoint URLs.
    .setAuthorizationBaseUrl(
      'https://www.linkedin.com/uas/oauth2/authorization')
    .setTokenUrl('https://www.linkedin.com/uas/oauth2/accessToken')

    // Set the client ID and secret.
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)

    // Set the scope of the request.
    .setScope('profile openid w_member_social')

    // Set the name of the callback function that should be invoked to
    // complete the OAuth flow.
    .setCallbackFunction('authCallback')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties());
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService_();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}
