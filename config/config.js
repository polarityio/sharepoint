module.exports = {
  /**
   * Name of the integration which is displayed in the Polarity integrations user interface
   *
   * @type String
   * @required
   */
  name: 'Sharepoint',
  /**
   * The acronym that appears in the notification window when information from this integration
   * is displayed.  Note that the acronym is included as part of each "tag" in the summary information
   * for the integration.  As a result, it is best to keep it to 4 or less characters.  The casing used
   * here will be carried forward into the notification window.
   *
   * @type String
   * @required
   */
  acronym: 'SHP',
  defaultColor: 'light-gray',
  onDemandOnly: true,
  /**
   * Description for this integration which is displayed in the Polarity integrations user interface
   *
   * @type String
   * @optional
   */
  description: 'Polarity Sharepoint integration',
  entityTypes: ['IP', 'hash', 'domain', 'string', 'email'],
  /**
   * An array of style files (css or less) that will be included for your integration. Any styles specified in
   * the below files can be used in your custom template.
   *
   * @type Array
   * @optional
   */
  styles: ['./styles/sharepoint.less'],
  /**
   * Provide custom component logic and template for rendering the integration details block.  If you do not
   * provide a custom template and/or component then the integration will display data as a table of key value
   * pairs.
   *
   * @type Object
   * @optional
   */
  block: {
    component: {
      file: './components/sharepoint-block.js'
    },
    template: {
      file: './templates/sharepoint-block.hbs'
    }
  },
  request: {
    // Provide the path to your certFile. Leave an empty string to ignore this option.
    // Relative paths are relative to the Sharepoint integration's root directory
    cert: '',
    // Provide the path to your private key. Leave an empty string to ignore this option.
    // Relative paths are relative to the Sharepoint integration's root directory
    key: '',
    // Provide the key passphrase if required.  Leave an empty string to ignore this option.
    // Relative paths are relative to the Sharepoint integration's root directory
    passphrase: '',
    // Provide the Certificate Authority. Leave an empty string to ignore this option.
    // Relative paths are relative to the Sharepoint integration's root directory
    ca: '',
    // An HTTP proxy to be used. Supports proxy Auth with Basic Auth, identical to support for
    // the url parameter (by embedding the auth info in the uri)
    proxy: ''
  },
  logging: {
    level: 'info' //trace, debug, info, warn, error, fatal
  },
  /**
   * Options that are displayed to the user/admin in the Polarity integration user-interface.  Should be structured
   * as an array of option objects.
   *
   * @type Array
   * @optional
   */
  options: [
    {
      key: 'host',
      name: 'Host',
      description:
        'The sharepoint host to use for querying data. This will typically look like `https://[TENANT-NAME].sharepoint.com`.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'clientId',
      name: 'Application (client) ID',
      description: 'The Application (client) ID to use for authentication.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'tenantId',
      name: 'Directory (tenant) ID',
      description: 'The Directory (tenant) id to authenticate inside of.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'privateKeyPath',
      name: 'Private Key File Path',
      description:
        'The Polarity Server file path to the private key file to use for authentication.  Relative paths should start with "./" and are relative to this integration\'s directory. The private key must be encoded in the PEM format using the PKCS8 container. Defaults to "./certs/private.key"',
      default: './certs/private.key',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'privateKeyPassphrase',
      name: 'Private Key Passphrase',
      description: 'The passphrase for the private key.  Leave blank if the private key does not have a passphrase.',
      default: '',
      type: 'password',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'publicKeyPath',
      name: 'Public Key File Path',
      description:
        'The Polarity Server file path to the public key file that corresponds to the private key used for authentication.  Relative paths should start with "./" and are relative to this integration\'s directory. The public key must be encoded in the PEM format using the PKCS8 container. Defaults to "./certs/public.key"',
      default: './certs/public.crt',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'subsite',
      name: 'Subsite Search Path',
      description:
        'Limit search to only the specified subsite path (optional).  Subsites can be specified by name, relative path, or full absolute path to include the host and scheme (e.g., https://[TENANT-NAME].sharepoint.com/sites/mysubsite)',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'blocklist',
      name: 'Ignore Entities',
      description: 'Comma delimited list of entities that you do not want to lookup.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: false
    },
    {
      key: 'domainBlocklistRegex',
      name: 'Ignore Domain Regex',
      description: 'Domains that match the given regex will not be looked up.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: false
    },
    {
      key: 'ipBlocklistRegex',
      name: 'Ignore IP Regex',
      description: 'IPs that match the given regex will not be looked up.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: false
    },
    {
      key: 'directSearch',
      name: 'Exact Match Search',
      description:
        'Check if you want each Sharepoint search to be an exact match with found entities (i.e., wrap the search term in quotes).',
      default: true,
      type: 'boolean',
      userCanEdit: false,
      adminOnly: true
    }
  ]
};
