// +------------------------------------------------------------------------------------------+
// |                                                                                          |
// | AzureADGraph.js                                                                          |
// | ---------------                                                                          |
// | Created by pinecat (https://github.com/pinecat/azure-ad-graph-expo)                      |
// |                                                                                          |
// | JavaScript library designed for use with Expo (https://expo.io).                         |
// | This library follows Microsoft's Azure AD authentication flow using                      |
// | Expo's own AuthSession to return user data from the Microsoft Graph                      |
// | API (/me endpoint).                                                                      |
// | You must register an app in Azure AD before you can authenticate using this method.      |
// | AzureADGraph is NOT a react component.                                                   |
// |                                                                                          |
// | https://docs.microsoft.com/en-us/azure/active-directory/develop/v1-protocols-oauth-code  |
// | https://docs.microsoft.com/en-us/graph/use-the-api                                       |
// |                                                                                          |
// +------------------------------------------------------------------------------------------+

/* imports */
import { AuthSession } from 'expo'; // AuthSession: for opening the authorization URL

/* class: AzureADGraph */
export default class AzureADGraph {
    /*
      constructor
    */
    constructor(props) {
      this.props = props;
    }

    /*
      getGraphData
        starts the auth process and returns user data if successful
      params:   none
      returns:  this.state.graphResponse - the user data from the MS Graph API
    */
    getGraphData() {
      return this.openAuthSession();
    }

    /*
      openAuthSession
        opens the authorization url to the Azure AD app
      params:   none
      returns:  void
    */
    openAuthSession = async () => {
      let authResponse = AuthSession.startAsync({
        authUrl:
          `https://login.microsoftonline.com/${encodeURIComponent(this.props.tenantId)}/oauth2/authorize?client_id=${encodeURIComponent(this.props.clientId)}&response_type=code&redirect_uri=${encodeURIComponent(this.props.redirectUrl)}`,
      });
      return this.getToken(authResponse.params.code);
    }

    /*
      getToken
        uses the code from the authResponse to retrieve a token
        this token can be used to call the Graph API
      params:   code - code retrieved from Azure AD auth
      returns:  void
    */
    getToken(code) {
      /* gather all required body params in an object */
      var requestParams = {
        clientId: this.props.clientId,
        scope: this.props.scope,
        code: code,
        redirectUrl: this.props.redirectUrl,
        grantType: 'authorization_code',
        clientSecret: this.props.clientSecret,
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
      fetch(`https://login.microsoftonline.com/${encodeURIComponent(this.props.tenantId)}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
        },
        body: formBody,
      })
      .then((response) => response.json())
      .then((tokenResponse) => {
        return this.callMsGraph(tokenResponse.access_token);
      })
      .catch((error) => {
        console.error(error);
      });
    }

    /*
      callMsGraph
        uses the token retrieved from the azure oauth token endpoint to get user data from the MS Graph API
      params:   token - unique token used to query Graph API
      returns:  void
    */
    callMsGraph(token) {
      /* make a GET request using fetch and querying with the token */
      fetch('https://graph.microsoft.com/v1.0/me', {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + token,
        }
      })
      .then((response) => response.json())
      .then((graphResponse) => {
        return graphResponse;
      })
      .catch((error) => {
        console.error(error);
      });
    }
}
