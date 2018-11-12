# PSGrantmaker

The PSGrantmaker PowerShell module can be used to interact with data using the Fluxx Grantmaker APIs. The calls are outlined below: 

### Get-FluxxBearerToken
Every call to the Fluxx API requires an authentication token. This call returns an object containing an token (access_token), token type (token_type), and the number of seconds before the token expires (expires_in). This should be the first call made before any others. The access_token attribute is used as the -BearerToken parameter in all of the other calls.

Before making this call, an OAuth application needs to be registered at **https://<your site>.fluxx.io/oauth/applications**. The person registering the application will need to be logged in as an admin before visiting the site. The Application ID and Secret provided during the registration are used for this call.

**Parameters**
  - **BaseURL**: The URL for the Fluxx instance being called.
  - **ApplicationID**: The application ID assigned when registering an OAuth application.
  - **Secret**: The secret assigned when registering an OAuth application.

**Example**
```sh
PS C:\Users\you> $baseUrl = "<your site>.fluxx.io"
PS C:\Users\you> $appId = "<your application id>"
PS C:\Users\you> $secret = "<your application secret>"
PS C:\Users\you> $token = Get-FluxxBearerToken -baseUrl $baseUrl -applicationId $appId -secret $secret
PS C:\Users\you> $token
access_token                                                     token_type expires_in
------------                                                     ---------- ----------
kas2q1jfd6izai4ovs36nta27i3bja1j0chi372qvpxl7kgya1ehnuv6yydaxqgq bearer           7200
```
