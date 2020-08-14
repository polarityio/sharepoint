# Polarity sharepoint Integration

The Polarity Sharepoint integration allows freeform text searching for IPs, Hashes and domains in your Sharepoint instance and retrieves related documents.

For more information on Sharepoint, please visit: [official website] (https://products.office.com/en-us/sharepoint/collaboration).

Check out the integration in action:

![sharepoint](https://user-images.githubusercontent.com/22529325/55797620-ed0c9900-5a9a-11e9-8438-b9ea09136081.gif)


## Sharepoint Integration Options

### Host

The top level domain of the Sharepoint instance to use for queries.

### Authentication Host

The host to use for authenticating the user, usually this should be allowed to default.

### Client ID

The Client ID to use for authentication.

### Client Secret

The Client Secret associated with the Client ID.

### Tenant ID

The Tenant ID of the Sharepoint instance.

### Subsite

Limit search to only a subsite.  This field should be only the subsite name, _not_ the full path.  This field is optional and can be left blank.

### Ingore List

Comma delimited list of domains that you do not want to lookup.

### Ignore Domain Regex

Domains that match the given regex will not be looked up..

### Ignore IP Regex

IPs that match the given regex will not be looked up.

## Polarity

Polarity is a memory-augmentation platform that improves and accelerates analyst decision making.  For more information about the Polarity platform please see:

https://polarity.io/
