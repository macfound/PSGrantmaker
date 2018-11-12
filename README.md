# PSGrantmaker

The PSGrantmaker PowerShell module can be used to interact with data using the Fluxx Grantmaker APIs. The calls are outlined below. To use this module, download PSGrantmaker.psm1 to your local file system (e.g. C:\Modules) and import the module:
```sh
PS C:\Users\you> Import-Module "C:\Modules\PSGrantmaker.psm1"
```

### Get-FluxxBearerToken
Every call to the Fluxx API requires an authentication token. This call returns an object containing an token (access_token), token type (token_type), and the number of seconds before the token expires (expires_in). This should be the first call made before any others. The access_token attribute is used as the -BearerToken parameter in all of the other calls.

Before making this call, an OAuth application needs to be registered at **https://<your site>.fluxx.io/oauth/applications**. The person registering the application will need to be logged in as an admin before visiting the site. The Application ID and Secret provided during the registration are used for this call.

**Parameters**
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***ApplicationID***: The application ID assigned when registering an OAuth application.
  - ***Secret***: The secret assigned when registering an OAuth application.

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

### Get-FluxxObject
Fluxx Grantmaker is made up of different objects. Some of the key objects include _grant_request_, _organziation_, and _user_. All objects are available via the API. This PowerShell function retrieves an object or list of objects via the Fluxx API and returns a serialized PSObject. The maximum number of records returned is 500 per page. The default is 25 per page.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be returned via the API (e.g. grant_request).
  - ***QuerystringParameters***: [Optional] Used to override and add querystring parameters. It should begin with the parameters used to define which columns to return. If filters need to be applied, use this parameter. Defaults to "all_core=1&all_dynamic=1"
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
 Rertieve the first 25 grant_request records via the API. Returns a PSbject including an array of records
```sh
PS C:\Users\you> Get-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "grant_request"
```
 
**Example**
Rertieve a specific organization record via the API. Returns a PSObject for a specific organization
```sh
PS C:\Users\you> Get-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "organization/<org id>"
```

**Example**
Rertieve the core fields of the first 100 request_transaction records due within the next 3 months that haven't been paid. Returns an object including an array of records
```sh
PS C:\Users\you> Get-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "request_transaction" -QuerystringParameters 'all_core=1&per_page=100&filter={"group_type":"and","conditions":[["due_at","next-n-months","3"],["paid_at","null","-"]]}'
```

### New-FluxxObject
This function is used to create a new object.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be created via the API (e.g. grant_request).
  - ***Data***: The data to be used to create the new object made up of name and value pairs using the following syntax:
  $subprogram = @{
    name = 'A Sub Program'
    description = 'The Sub Program Description'
    program_id = 1}
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
Create a new sub_program record via the API. Returns the newly created sub_program PSObject
```sh
PS C:\Users\you> $subprogram = @{
                     name = 'A Sub Program'
                     description = 'The Sub Program Description'
                     program_id = 1
                 }
PS C:\Users\you> New-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "sub_program" -Data $subprogram
```

### Set-FluxxObject
This function is used to update an existing object.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be updated via the API (e.g. grant_request).
  - ***RecordID***: The ID of the record to be updated. The value should be an integer.
  - ***Data***: The data to be used to update the existing object made up of name and value pairs using the following syntax:
  $program = @{
    name = 'An Existing Program'
    description = 'Updated Program Description'
  }
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
Update an existing new sub_program record via the API. Returns the newly created PSObject
```sh
PS C:\Users\you> $subprogram = @{
                     name = 'An Existing Sub Program'
                     description = 'The Updated Sub Program Description'
                 }
PS C:\Users\you> Set-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "sub_program" -RecordID 12345 -Data $subprogram
```

### Remove-FluxxObject
This function is used to delete an existing object.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be deleted via the API (e.g. grant_request).
  - ***RecordID***: The ID of the record to be deleted. The value should be an integer.
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
Delete an existing sub_program record with the ID 12345. Returns TRUE or FALSE depending on success
```sh
PS C:\Users\you> Remove-FluxxObject -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "sub_program" -RecordID 12345
```

### Get-FluxxObjectList
This function leverages the Get-FluxObject function to return a list of all records. The Get-FluxxObject function can return up to 500 records. This function identifies the number of records and continually makes calls until all records are returned. A list of objects is returned.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be returned via the API (e.g. grant_request).
  - ***RecordsPerPage***: [Optional] The nuber of records to be returned per page. The default is 100 but can be increased up to 500. A larger number reduces the number of calls required to retrieve the full list but can have performance impacts on your tenant when calling objects with a large number of attributes.
  - ***QuerystringParameters***: [Optional] Used to override and add querystring parameters. It should begin with the parameters used to define which columns to return. If filters need to be applied, use this parameter. Defaults to "all_core=1&all_dynamic=1"
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
Retrieve all initiatives stored within Fluxx pulling 500 records at a time.
```sh
PS C:\Users\you> Get-FluxxObjectList -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "initiative" -RecordsPerPage 500"
```

**Example**
Retrieve all grant_request records stored within Fluxx without a workflow state of declined.
```sh
PS C:\Users\you> Get-FluxxObjectList -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "grant_request" -QueryStringParameters 'filter={"group_type":"or","conditions":[["state","not-eq","declined"]]}'
```

### Import-FluxxDocument
This function is used to upload a document into Fluxx Grantmaker.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be returned via the API.  Defaults to "model_document".
  - ***FileName***: The full path of the file to be uploaded into Fluxx (e.g. C:\files\file-to-upload.txt).
  - ***ModelType***: The model type of the record associated with this file (e.g. grant_request).
  - ***ModelTypeOwnerID***: The ID of the model type of the record associated with this file.
  - ***ModelTypeID***: The ID of the model_document_type associated with the file to be uploaded.
  - ***UserID***: The ID of the people record associated with the upload.
  - ***Description***: A description of the file being uploaded.
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
Import a document via the Fluxx API.
```sh
PS C:\Users\you> Import-FluxxDocument -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -FileName "C:\files\file-to-upload.txt" -ModelType "request_report" -ModelTypeOwnerID 23213 -ModelTypeID 55313 -UserID 41413 -Description "Annual report summary"
```

### Export-FluxxDocument
This function is used to download a document that has been uploaded into Fluxx Grantmaker.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be returned via the API.  Defaults to "model_document".
  - ***DocumentID***: The ID of the document to be retrieved. The value should be an integer.
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"

**Example**
Retrieve a document with the ID of 1026325 via the Fluxx API.
```sh
PS C:\Users\you> Export-FluxxDocument -BearerToken <bearer token> -BaseUrl "<your site>.fluxx.io" -DocumentID 1026325
```

### Export-FluxxObjectListToSQLServer 
This function leverages the Get-FluxObjectList function to return a list of all records for a given Fluxx object and pushes it into a table within a Microsoft SQL Server database.

**Parameters**
  - ***BearerToken***: A token used to authenticate against the Fluxx API. Use the Get-FluxxBearerToken function and reference the access_token attribute of the retrieved PSObject.
  - ***BaseURL***: The URL for the Fluxx instance being called.
  - ***FluxxObject***: The name of the object type to be returned via the API (e.g. grant_request).
  - ***RecordsPerPage***: [Optional] The nuber of records to be returned per page. The default is 100 but can be increased up to 500. A larger number reduces the number of calls required to retrieve the full list but can have performance impacts on your tenant when calling objects with a large number of attributes.
  - ***QuerystringParameters***: [Optional] Used to override and add querystring parameters. It should begin with the parameters used to define which columns to return. If filters need to be applied, use this parameter. Defaults to "all_core=1&all_dynamic=1"
  - ***ApiVersion***: [Optional] Allows for overriding the version of the API to use. Defaults to "v2"
  - ***SQLServerName***: The fully qualified domain name of a SQL Server instance.
  - ***SQLDatabaseName***: The name of the database in which the table will be created.
  - ***SQLUserName***: The username to be used to connect to the SQL Server.
  - ***SQLPassword***: The password to be used to connect to the SQL Server.
  - ***SQLSchema***: [Optional] A schema name to be used. Defaults to fluxx
  - ***OverwriteTable***: [Optional] A switch used to to drop and recreate the table. If the switch is not used, records will be appended to an existing table if one exists.

**Example**
Retrieve all initiatives stored within Fluxx pulling 500 records at a time and automatically create a clean table within a MS SQL Server.
```sh
PS C:\Users\you> Export-FluxxObjectListToSQLServer -BearerToken  <bearer token> -BaseUrl "<your site>.fluxx.io" -FluxxObject "grant_request" -SQLServerName <sql server name> -SQLDatabaseName <database name> -SQLUserName <database user name> -SQLPassword <database password> -OverwriteTable
```
