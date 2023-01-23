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
  entityTypes: ['ip', 'hash', 'domain', 'string', 'email'],
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
    proxy: '',
    /**
     * If set to false, the integration will ignore SSL errors.  This will allow the integration to connect
     * to Sharepoint servers without valid SSL certificates.  Please note that we do NOT recommending setting this
     * to false in a production environment.
     */
    rejectUnauthorized: true
  },
  logging: {
    // directory is relative to the this integrations directory
    // e.g., if the integration is in /app/polarity-server/integrations/virustotal
    // and you set directoryPath to be `integration-logs` then your logs will go to
    // `/app/polarity-server/integrations/integration-logs`
    // You can also set an absolute path.  If you set an absolute path you must ensure that
    // the directory you specify is writable by the `polarityd:polarityd` user and group.

    //directoryPath: '/var/log/polarity-integrations',
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
      key: 'authHost',
      name: 'Authentication Host',
      description:
        'The authentication host to use for querying data.  This should usually be set to the default value of "https://accounts.accesscontrol.windows.net".',
      default: 'https://accounts.accesscontrol.windows.net',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'clientId',
      name: 'Client ID',
      description: 'The client ID to use for authentication.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'clientSecret',
      name: 'Client Secret',
      description: 'The secret to use for authentication.',
      default: '',
      type: 'password',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'tenantId',
      name: 'Tenant ID',
      description: 'The tenant id to authenticate inside of.',
      default: '',
      type: 'text',
      userCanEdit: false,
      adminOnly: true
    },
    {
      key: 'subsite',
      name: 'Subsite',
      description:
        'Limit search to only a subsite (optional).  This field should be only the subsite name and not the full path.',
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
      name: 'Direct Search',
      description:
        'Check if you want each Sharepoint search to be an exact match with found entities (i.e., wrap the search term in quotes).',
      default: false,
      type: 'boolean',
      userCanEdit: false,
      adminOnly: true
    }
  ]
};
