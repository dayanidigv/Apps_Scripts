function fetchLinkedInProfile(accessToken) {
  const service = getService_();
  if (!service.hasAccess()) {
    Logger.log('Authorization required.');
    return;
  }

  var url = 'https://api.linkedin.com/v2/me';

  var options = {
    'headers': {
      Authorization: 'Bearer ' + service.getAccessToken()
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  // Output the person's URN
  Logger.log('URN: ' + json.id);  // This is the URN you're looking for
}


function postToLinkedIn() {
  const service = getService_();
  if (!service.hasAccess()) {
    Logger.log('Authorization required.');
    return;
  }

  const url = 'https://api.linkedin.com/v2/posts';
  const payload = {
    author : 'urn:li:person:XXXX',
    commentary : "Hello, these are some bullet points:\n\n\\* Point 1\n\\* Point 2\n\\* Point 3",
    visibility : "PUBLIC",
    distribution : {
      feedDistribution : "MAIN_FEED",
      targetEntities: [],
      thirdPartyDistributionChannels: []
    },
    lifecycleState: "PUBLISHED",
    isReshareDisabledByAuthor: false
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken()
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

function sharesToLinkedIn() {
  const service = getService_();
  if (!service.hasAccess()) {
    Logger.log('Authorization required.');
    return;
  }

  const url = 'https://api.linkedin.com/v2/shares';
  const payload = {
    owner: 'urn:li:person:XXXX',
    subject: 'Automation',
    text: {
      text: "Hello, these are some bullet points:\n\n\\* Point 1\n\\* Point 2\n\\* Point 3"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken()
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}




