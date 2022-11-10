# Polarity sharepoint Integration

The Polarity Sharepoint integration allows freeform text searching for IPs, Hashes and domains in your Sharepoint instance and retrieves related documents.

For more information on Sharepoint, please visit: [official website] (https://products.office.com/en-us/sharepoint/collaboration).

Check out the integration in action:

![sharepoint](https://user-images.githubusercontent.com/22529325/55797620-ed0c9900-5a9a-11e9-8438-b9ea09136081.gif)

## Configuring Sharepoint

The Polarity-Sharepoint integration uses Sharepoint Addin Authentication via OAuth bearer tokens.  To setup the integration you register a new application with Sharepoint to generate the client id and client secret.  Once the application is registered you set the permissions.

### Register the App

Navigate to `https://[TENANT-NAME].sharepoint.com/_layouts/15/appregnew.aspx`.  Click on "Generate" for the `Client ID` and record the value.
Click on generate for the `Client Secret` and record the value. 

Fill in a `Title` such as "Polarity Sharepoint Integration".

For the AppDomain you can specify your company domain (e.g., mycompany.com) and for the Redirect URI you can use `https://localhost/`.

Click Create.

### Give permissions

Navigate to `https://[TENANT-NAME]-admin.sharepoint.com/_layouts/15/appinv.aspx`. 
> Note that this will register the application at the tenant level which means the credentials can be used everywhere inside your tenant.

Fill in the `clientId` as the `App Id` and click on the "Lookup" button.

Once the inputs fill in with your app information provide the following permission.

```
<AppPermissionRequests AllowAppOnlyPolicy="true">
  <AppPermissionRequest Scope="http://sharepoint/content/tenant" Right="Read" />
</AppPermissionRequests>
```

This will provide `READ` access to sites within your tenant.

Click "Create" and then the "Trust It" button.

### Finding your TenantId

To find your TenantId navigate to the url `https://[TENANT-NAME].sharepoint.com/_layouts/appprincipals.aspx`.  When the page loads, in the right hand `App Identifier` column you will find a series of app ids.  The Tenant ID appears after the `@` sign and has the format `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`.

## Sharepoint Integration Options

### Host

The sharepoint host to use for querying data.  This will typically look like `https://[TENANT-NAME].sharepoint.com`.

### Authentication Host

The authentication host to use for querying data.  This should usually be set to the default value of `https://accounts.accesscontrol.windows.net`.

### Client ID

The Client ID to use for authentication.

### Client Secret

The Client Secret associated with the Client ID.

### Tenant ID

The Tenant ID to authenticate inside of.

### Subsite

Limit search to only a subsite.  This field should be only the subsite name, _not_ the full path.  This field is optional and can be left blank.

### Ignore Entities

Comma delimited list of entities that you do not want to lookup.

### Ignore Domain Regex

Domains that match the given regex will not be looked up.

### Ignore IP Regex

IPs that match the given regex will not be looked up.

## Polarity

Polarity is a memory-augmentation platform that improves and accelerates analyst decision making.  For more information about the Polarity platform please see:

https://polarity.io/
