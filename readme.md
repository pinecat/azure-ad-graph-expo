# azure-ad-graph-expo
#### by pinecat

### About
This is a simple JavaScript library designed to be used with [Expo](https://expo.io).  This utilizes Expo's AuthSession to authenticate via Microsoft Azure AD.  It follows Microsoft's Azure authentication flow to first login the user, then acquire a token, and then use that token to query the Microsoft Graph API /me endpoint to get user data.

### Azure Endpoints & Expo
This library now uses the Azure v2 endpoints!  If you must use the v1 endpoints, please use version 1.1.2.  Also, version 1 of this library was originally intended for use with Expo v35.  This version works with Expo v37, the latest version of Expo to date.

### Installing
You can install this library via `npm` or `yarn` like so:
```sh
$ npm install azure-ad-graph-expo
```
```sh
$ yarn add azure-ad-graph-expo
```

### Example Code
```javascript
import React from 'react';
import { StyleSheet, View, Text, Button } from 'react-native'
import * as AuthSession from 'expo-auth-session';
import { openAuthSession } from 'azure-ad-graph-expo';

export default class App extends React.Component {
  state = {
    result: null,
  };

  render() {
    return (
      <View style={styles.container}>
        <Button title="Login" onPress={this._handlePressAsync} />
        {this.state.result ? (
          <Text>{JSON.stringify(this.state.result)}</Text>
        ) : <Text>Nothing to see here.</Text>}
      </View>
    );
  }

  _handlePressAsync = async () => {
    let result = await openAuthSession(azureAdAppProps);
    this.setState({ result });
  }
}

const azureAdAppProps = {
        clientId        :   AZURE_CLIENT_ID,
        tenantId        :   AZURE_TENANT_ID,
        scope           :   'user.read',
        redirectUrl     :   AuthSession.makeRedirectUri(),
        clientSecret    :   AZURE_CLIENT_SECRET,
        domainHint      :   AZURE_DOMAIN_HINT,
        prompt          :   'login'
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#fff',
    alignItems: 'center',
    justifyContent: 'center',
  },
});
```
