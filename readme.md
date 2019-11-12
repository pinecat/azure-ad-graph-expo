# azure-ad-graph-expo
#### by pinecat

### About
This is a simple JavaScript library designed to be used with [Expo](https://expo.io).  This utilizes Expo's AuthSession to authenticate via Microsoft Azure AD.  It follows Microsoft's Azure authentication flow to first login the user, then acquire a token, and then use that token to query the Microsoft Graph API /me endpoint to get user data.

### Example Code
```javascript
import React from 'react';
import { StyleSheet, View, Text, Button } from 'react-native'
import { AuthSession } from 'expo';
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
  clientId: 'client_id',
  tenantId: 'tenant_id',
  scope: 'user.read',
  redirectUrl: AuthSession.getRedirectUrl(),
  clientSecret: 'client_secret',
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
