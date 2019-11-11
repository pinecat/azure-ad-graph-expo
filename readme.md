# azure-ad-graph-expo
#### by pinecat

### About
This is a simple JavaScript library designed to be used with [Expo](https://expo.io).  This utilizes Expo's AuthSession to authenticate via Microsoft Azure AD.  It follows Microsoft's Azure authentication flow to first login the user, then acquire a token, and then use that token to query the Microsoft Graph API /me endpoint to get user data.

### Example Code
```javascript
import React from 'react';
import { View, Text, Button, AuthSession } from 'expo';

export default class App extends React.Component {

  render() {
    return (
      <View>
        
      </View>
    );
  }
}
```
