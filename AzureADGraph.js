// +------------------------------------------------------------------------------------------+
// |                                                                                          |
// | AzureADGraph.js                                                                          |
// | ---------------                                                                          |
// | Created by pinecat (https://github.com/pinecat/azure-ad-graph-expo)                      |
// |                                                                                          |
// | CONTRIBUTORS:                                                                            |
// |  pinecat (https://github.com/pinecat)                                                    |
// |  JuanDavidLopez95 (https://github.com/JuanDavidLopez95)                                  |
// |                                                                                          |
// | JavaScript library designed for use with Expo (https://expo.io).                         |
// | This library follows Microsoft's Azure AD authentication flow using                      |
// | Expo's own AuthSession to return user data from the Microsoft Graph                      |
// | API (/me endpoint).                                                                      |
// | You must register an app in Azure AD before you can authenticate using this method.      |
// | AzureADGraph is NOT a react component.                                                   |
// |                                                                                          |
// | https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-overview              |
// | https://docs.microsoft.com/en-us/graph/use-the-api                                       |
// |                                                                                          |
// +------------------------------------------------------------------------------------------+

/* imports */
import * as AuthSession from 'expo-auth-session'; // AuthSession: for opening the authorization URL

/*
    callMsGraph
      queries the Microsoft Graph API to return user data
      params:   token - the unique token used to query the logged in user in the Graph API
      returns:  graphResponse - a JSON object of the user data
*/
const callMsGraph = async (token) => {
  /* make a GET request using fetch and querying with the token */
  let graphResponse = null;
  await fetch('https://graph.microsoft.com/v1.0/me', {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + token,
    }
  })
  .then((response) => response.json())
  .then((response) => {
    graphResponse = response;
  })
  .catch((error) => {
    graphResponse = error;
  });
  
  /* 
    Spread the results of the graph and add a type property with a value of success to indicate
    that the AzureAD info grabbing was a success
  */ 
  const finalResponse = {
    ...graphResponse,
    type: "success"
  }

  return finalResponse;
}; //end callMsGraph()

/*
  getToken
        sends POST request to MS Azure AD token endpoint to get a token that can be used to
        query the MS Graph API.  also forms the body to be send in the POST request
        params:   code - the code that was recieved when the user logged in via Azure
                props - the Azure AD app properties used to construct the request
        returns:  callMsGraph() - callMsGraph will return the user data from the Graph API
*/
const getToken = async (code, props) => {
    /* parse/gather correct key values for the POST request to the token endpoint */
    /* Client secret is omitted here; including it yields an error */
    var requestParams = {
      client_id: props.clientId,
      scope: props.scope,
      code: code,
      redirect_uri: props.redirectUrl,
      grant_type: 'authorization_code',
    }
  
    /* loop through object and encode each item as URI component before storing in array */
    /* then join each element on & */
    /* request is x-www-form-urlencoded as per docs: https://docs.microsoft.com/en-us/graph/use-the-api */
    var formBody = [];
    for (var p in requestParams) {
      var encodedKey = encodeURIComponent(p);
      var encodedValue = encodeURIComponent(requestParams[p]);
      formBody.push(encodedKey + '=' + encodedValue);
    }
    formBody = formBody.join('&');
  
    /* make a POST request using fetch and the body params we just setup */
    let tokenResponse = null;
    await fetch(`https://login.microsoftonline.com/${props.tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
      },
      body: formBody,
    })
    .then((response) => response.json())
    .then((response) => {
      tokenResponse = response;
    })
    .catch((error) => {
      console.error(error);
    });
    return await callMsGraph(tokenResponse.access_token);
}; //end getToken()

/*
  openAuthSession
    opens connection to the azure ad authentication url using Expo's AuthSession
    params:   props - a JSON object of your app properies (i.e. clientId, tenantId, etc)
    returns:  getToken() - calls getToken which calls callMsGraph to return the user data from the Graph API
*/
const openAuthSession = async (props) => {
  const authUrl = `https://login.microsoftonline.com/${props.tenantId}/oauth2/v2.0/authorize?client_id=${props.clientId}&response_type=code&scope=${encodeURIComponent(props.scope)}${props.domainHint ? "&domain_hint=" + encodeURIComponent(props.domainHint) : null}${props.prompt ? "&prompt=" + props.prompt : null}&redirect_uri=${encodeURIComponent(props.redirectUrl)}`;

  /*
    Add a returnUrl parameter to the AuthSession.startAsync() config object.
    If undefined, fall back to AuthSession.makeRedirectUri().
    Strange bug on Android --
    Without this, even if authentication is successful, the result of authResponse will always 
    be { type: 'dismiss' }

    Source :
        https://docs.expo.io/versions/latest/sdk/auth-session/#authsessionstartasyncoptions
  */ 
  
  let authResponse = await AuthSession.startAsync({
      authUrl     :   authUrl,
      returnUrl   :   props.returnUrl || AuthSession.makeRedirectUri()
    })
    .then((authResponse ) => { 
          //Conditional if the user proceeds with the authentication process
          if (authResponse.type === "success") { 

          /*
            Only continue with the authentication process if user does not cancel or close 
            the ongoing authentication window or session. authResponse and the code from 
            the parameters will be defined if the authorization session continues. 
          */
          
            //Do not proceed with acquiring a token if there is an error. Return the error and the response.
            if (authResponse.params["error"] ) {

              /* 
                Return error as an added error property value in the object for easier 
                catching of the error in the front end 
              */
              return { 
                      "error": authResponse.params.error,
                      ...authResponse
                    };
            } else {
              //If authentication is successful, pass the authorization code to get the token
              return getToken(authResponse.params.code, props);
            }

          } //end if-statement
          //Else statement to catch if the user has decline continuing further with authentication.
          else { 
            //Return the authResponse which wil include a type of dismissed or cancelled. 
            return { 
              "error": "Authorization session cancelled",
              ...authResponse
            };
          } //end else-statement.
    })
    .catch((error) => {
      console.error(error);
      return { 
          "error" : error , 
          "type": "error"
        };
    });

    //Return the final authResponse that resolves from the promise.
    return authResponse;
}; //end openAuthSession()

export { openAuthSession };