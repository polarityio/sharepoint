# Polarity sharepoint Integration

The Polarity Sharepoint integration allows searching of IPs, Hashes, domains, cves, emails and annotated entities in your Sharepoint instance and retrieves related documents.

For more information on Sharepoint, please visit: [official website] (https://products.office.com/en-us/sharepoint/collaboration).

Check out the integration in action:

![sharepoint](https://user-images.githubusercontent.com/22529325/55797620-ed0c9900-5a9a-11e9-8438-b9ea09136081.gif)

## Upgrade Notes from v3.4.1

Versions of this integration up to v3.4.1 supported authentication as a Sharepoint Add-in using a Client Secret.  This method of authentication is discouraged by Microsoft and is no longer supported by this version of the integration.  If you are using a previous version of the integration you will need to reconfigure the integration to use Azure App Authentication via Certificates.  See the "Configuring Sharepoint" section below for more information.

## Network Connectivity

Note that to authenticate with Sharepoint in Azure, the integration will need access to both your Sharepoint site (e.g., https://mysite.sharepoint.com) and it will need access to the Microsoft authentication server
`https://login.microsoftonline.com`.

If your environment requires configuring a proxy, ensure that the proxy allows outbound connections to both your sharepoint site and `https://login.microsoftonline.com`.

## Configuring Sharepoint

The Polarity-Sharepoint integration uses Azure App Authentication via Certificates.  To configure the integration you register a new application with Azure.  Once the application is registered you will need to upload a public certificate (the corresponding private certificate is needed by the integration on the Polarity Server).  Finally, you need to set the appropriate API permissions.  See below for detailed instructions. 

### Register the App

Browse to https://entra.microsoft.com/ and then navigate to "Applications" -> "App Registrations".  From there, click on "New registration" to create a new Azure Application.

Choose a name for your application (e.g., "Polarity Sharepoint Integration").  A Redirect URI is not required for this integration.

Once the application is created, record the `Application (client) ID` and `Directory (tenant) ID` values as you will need these when configuring the integration in Polarity.

<img src="./images/app-id-tenant-id.png" width="500px">

### Upload a public key certificate

Sharepoint Azure Apps require that  app-based authentication is done via a certificate (Client secret based authentication is not supported by Azure for Sharepoint).  Upload your public key to Azure by navigating to "Certificates & secrets" and clicking on the "Upload certificate" button to upload your **public** key to Azure. 

<img src="./images/upload-cert.png" width="500px">

The public/private key pair must be encoded in the **PEM** format using the **PKCS8** container.  

You can check to see if your private key is in the right format by viewing the content of the key.  The file content should look like this:

```
-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC1NALrZ6xcgCLO
// Additional Base64 encoded text
-----END PRIVATE KEY-----
```

You can create a self-signed public/private key pair using the following command:

```
openssl req -x509 -nodes -sha256 -days 365 -newkey rsa:2048 -keyout private.key -out public.crt
```

> Note that the above command will create a private key that is not encrypted with a passphrase.  If you want to encrypt the private key with a passphrase you can add the `-passout pass:<your-password>` option to the command above.

The above command will generate a private key called `private.key` and a public key called `public.crt`.  Upload `public.crt` to your Azure application. 

Both the public and private key should be saved on the Polarity Server.  By default, the integration will look for the public and private key in the `./certs` directory.  To use the default settings, your directory structure should look like this:

```
| /app/polarity-server/integration/sharepoint/certs
|------- public.crt
|------- private.key
```

> The public key file is required so that the integration can compute the thumbprint of the public key.  The thumbprint is used to authenticate with Azure.

If you need to change the location or filename of the public and private key be sure to update the "Private Key File Path" and "Public Key File Path" integration options.


### Set API Permissions

From the Azure "App Registrations" page, click on "API Permissions" and then click on the "Add a permission" button.  

Under Microsoft APIs select "SharePoint" and then select "Application permissions".

Under "Sites", find the permissions for "Sites.Read.All" and select it.  Click on the "Add permissions" button to add the permission.

Once the permission has been granted you will need to make sure to "Grant admin consent" for the permission.

<img src="./images/api-permissions.png" width="500px">

In this example we granted permission via the "Sites.Read.All" permission.  You can further constrain site access via the "Subsite Search Path" option within the Polarity Integration.

The integration also supports the SharePoint "Sites.Selected" permission.  This permission allows you to specify specific sites that the integration can access.  This access is enforced by Azure.  If you use this permission, you will also need to grant the appropraite permissions to the Azure app via the Graph API permissions endpoint.  

For more information on the "Sites.Selected" permission please see https://devblogs.microsoft.com/microsoft365dev/updates-on-controlling-app-specific-access-on-specific-sharepoint-sites-sites-selected/


### Configure the Integration

1. SSH into the Polarity Server and copy or upload the public and private keys into the integration's `certs` directory.  We recommend that your keys are named `private.key` and `public.crt`.  

2. Ensure the keys are readable by the `polarityd` user:

```
sudo chown polarityd:polarityd /app/polarity-server/integrations/sharepoint/certs/private.key
sudo chown polarityd:polarityd /app/polarity-server/integrations/sharepoint/certs/public.key
```

3. From the integration's settings page (within Polarity), fill in your SharePoint host information which will typically look like `https://[TENANT-NAME].sharepoint.com`.

4. Fill in the Application (client ID), and Directory (tenant) ID values that you recorded when you registered your Azure application.

5. If you placed your `private.key` and `public.crt` files into the default locations you do not need to modify the "Private/Public Key File Path" options.

6. If your private key has a passphrase then provide the passphrase for the "Private Key Passphrase" option.

The integration should now be able to authenticate to your SharePoint instance.  You can further configure the integration by reviewing the full options list below.

## Sharepoint Integration Options

### Host
The sharepoint host to use for querying data. This will typically look like `https://[TENANT-NAME].sharepoint.com`.

### Application (client) ID

The Application (client) ID to use for authentication.

### Directory (tenant) ID

The Directory (tenant) id to authenticate inside of.

### Private Key File Path

The Polarity Server file path to the private key file to use for authentication. Relative paths should start with "./" and are relative to this integration's directory. The private key must be encoded in the PEM format using the PKCS8 container. 

Defaults to "./certs/private.key".

You must restart the integration after making changes to this option.

### Private Key Passphrase

The passphrase for the private key. Leave blank if the private key does not have a passphrase.

### Public Key File Path

The Polarity Server file path to the public key file that corresponds to the private key used for authentication. Relative paths should start with "./" and are relative to this integration's directory. The public key must be encoded in the PEM format using the PKCS8 container. 

Defaults to "./certs/public.crt".

You must restart the integration after making changes to this option.

### Subsite Search Path

Limit search to only the specified subsite path (optional).  Subsites can be specified by name, relative path, or full absolute path to include the host.

#### Examples

The Subsite Search Path should mirror the URL you navigate to when viewing the site in question.  If the host for the Subsite matches the "Host" option value then you only need to provide the relative Subsite path. For example, if the URL path looks like this:

```
https://[tenant].sharepoint.com/mysubsite
```

Then you would set the "Subsite Search Path" value to:

```
mysubsite
```

In some cases the Subsite is proceeded by the word "sites".  For example:

```
https://[tenant].sharepoint.com/sites/mysubsite
```

In this scenario your "Subsite Search Path" would be set to:

```
sites/mysubsite
```

Finally, if the host for your Subsite differs from the Host integration option, then you should provide the entire absolute path to the Subsite.  This will happen for "personal" subsites which are located at `https://[tenant]-my`.  For example, if the Subsite is located at:

```
https://[tenant]-my.sharepoint.com/personal
```

Then you would set the "Subsite Search Path" value to:

```
https://[tenant]-my.sharepoint.com/personal
```

> Note that the Subsite name is not case-sensitive

### Ignore Entities
Comma delimited list of entities that you do not want to lookup.

### Ignore Domain Regex
Domains that match the given regex will not be looked up.

### Ignore IP Regex
IPs that match the given regex will not be looked up.

### Exact Match Search
Check if you want each Sharepoint search to be an exact match with found entities (i.e., wrap the search term in quotes).

## Troubleshooting

### Server timestamp

If you see the following error:

```
ServerError: invalid_client: 700024 - [2023-09-15 14:47:13Z]: AADSTS700024: Client assertion is not within its valid time range.
```

Then check to ensure that the clock on your server is set properly.

### Invalid API permissions

If you see any of the following errors then it's likely that you have not assigned the correct permissions to your Azure application.  Ensure the application has been assigned the `SharePoint -> Sites.Read.All` permission and that Admin consent has been granted.

```
"ID3035: The request was not valid or is malformed."
```

or

```
Error: status code was 401
```

### Missing Files

If you see an error like this:

```
Error: ENOENT: no such file or directory, open './certs/private.key'
```
It means the file could not be found.  Please check that the private or public key exists on your server in the expected location and that the file is readable by the `polarityd` user.


### 403 status code

A 403 status code means the credentials for the integration do not have permission to access SharePoint.  This typically occurs if the wrong public/private key pair is being used, or if the `Sites.Selected` permission is being used and the correct permissions have not been granted via the Graph API.

```
Error: status code was 403 at
```

### Permission Denied

If you see an error like this:

```
Error: EACCES: permission denied, open './certs/private.key'
```

It means the public and or private key are not readable by the `polarityd` user.  Ensure that the files are globally readable or are owned by the `polarityd` user.


### Proxy Issues

If the integration is unable to connect to your proxy or if the proxy is not allowing outbound connections to `login.microsoftonline.com` then you may see the following error:

```
ClientAuthError: endpoints_resolution_error: Error: could not resolve endpoints. Please check network and try again.
```

If you see this error, confirm that you have a proxy configured and that the configuration is correct.  You can also test for connectivity from the Polarity Server command line using curl:

```
curl -vvv https://login.microsoftonline.com
```

If you are using a proxy then add the proxy flag (`-x`) and your proxy configuration string:

```
curl -vvv -x https://myproxy:8080 https://login.microsoftonline.com
```

## Polarity

Polarity is a memory-augmentation platform that improves and accelerates analyst decision making.  For more information about the Polarity platform please see:

https://polarity.io/
