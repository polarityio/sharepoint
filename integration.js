const async = require('async');
const request = require('postman-request');
const util = require('util');
const url = require('url');
const config = require('./config/config.json');
const xbytes = require('xbytes');
const NodeCache = require('node-cache');
const msal = require('@azure/msal-node');
const crypto = require('crypto');
const SharepointHttpAuthClient = require('./sharepoint-http-client');

const cache = new NodeCache({
  stdTTL: 60 * 10
});

const fileTypes = {
  pdf: 'file-pdf',
  html: 'file-code',
  csv: 'file-csv',
  zip: 'file-archive',
  jpg: 'image',
  png: 'image',
  gif: 'image',
  xlsx: 'file-excel',
  docx: 'file-word',
  doc: 'file-word',
  ppt: 'file-powerpoint',
  pptx: 'file-powerpoint',
  html: 'file'
};

const MAX_PARALLEL_LOOKUPS = 10;
let Logger;
let requestWithDefaults;

let domainBlockList = [];
let previousDomainBlockListAsString = '';
let previousDomainRegexAsString = '';
let previousIpRegexAsString = '';
let domainBlocklistRegex = null;
let ipBlocklistRegex = null;
let clientApplication = null;

function _setupRegexBlocklists(options) {
  if (options.domainBlocklistRegex !== previousDomainRegexAsString && options.domainBlocklistRegex.length === 0) {
    Logger.debug('Removing Domain Blocklist Regex Filtering');
    previousDomainRegexAsString = '';
    domainBlocklistRegex = null;
  } else {
    if (options.domainBlocklistRegex !== previousDomainRegexAsString) {
      previousDomainRegexAsString = options.domainBlocklistRegex;
      Logger.debug({ domainBlocklistRegex: previousDomainRegexAsString }, 'Modifying Domain Blocklist Regex');
      domainBlocklistRegex = new RegExp(options.domainBlocklistRegex, 'i');
    }
  }

  if (options.blocklist !== previousDomainBlockListAsString && options.blocklist.length === 0) {
    Logger.debug('Removing Domain Blocklist Filtering');
    previousDomainBlockListAsString = '';
    domainBlockList = null;
  } else {
    if (options.blocklist !== previousDomainBlockListAsString) {
      previousDomainBlockListAsString = options.blocklist;
      Logger.debug({ domainBlocklist: previousDomainBlockListAsString }, 'Modifying Domain Blocklist Regex');
      domainBlockList = options.blocklist.split(',').map((item) => item.trim());
    }
  }

  if (options.ipBlocklistRegex !== previousIpRegexAsString && options.ipBlocklistRegex.length === 0) {
    Logger.debug('Removing IP Blocklist Regex Filtering');
    previousIpRegexAsString = '';
    ipBlocklistRegex = null;
  } else {
    if (options.ipBlocklistRegex !== previousIpRegexAsString) {
      previousIpRegexAsString = options.ipBlocklistRegex;
      Logger.debug({ ipBlocklistRegex: previousIpRegexAsString }, 'Modifying IP Blocklist Regex');
      ipBlocklistRegex = new RegExp(options.ipBlocklistRegex, 'i');
    }
  }
}

function _isEntityBlocklisted(entityObj, options) {
  if (domainBlockList.indexOf(entityObj.value) >= 0) {
    return true;
  }

  if (entityObj.isIPv4 && !entityObj.isPrivateIP) {
    if (ipBlocklistRegex !== null) {
      if (ipBlocklistRegex.test(entityObj.value)) {
        Logger.debug({ ip: entityObj.value }, 'Blocked BlockListed IP Lookup');
        return true;
      }
    }
  }

  if (entityObj.isDomain) {
    if (domainBlocklistRegex !== null) {
      if (domainBlocklistRegex.test(entityObj.value)) {
        Logger.debug({ domain: entityObj.value }, 'Blocked BlockListed Domain Lookup');
        return true;
      }
    }
  }

  return false;
}

function formatSearchResults(searchResults, options) {
  let data = [];

  searchResults.PrimaryQueryResult.RelevantResults.Table.Rows.forEach((row) => {
    let obj = {};

    row.Cells.forEach((cell) => {
      if (cell.Key === 'HitHighlightedSummary' && cell.Value) {
        cell.Value = cell.Value.replace(/c0/g, 'strong').replace(/<ddd\/>/g, '&#8230;');
      }

      obj[cell.Key] = cell.Value;
      if (cell.Value) {
        if (cell.Key === 'FileType') {
          if (fileTypes[cell.Value]) {
            obj._icon = fileTypes[cell.Value];
          } else {
            obj._icon = 'file';
          }
        }

        if (cell.Key === 'Size') {
          obj._sizeHumanReadable = xbytes(cell.Value);
        }

        if (cell.Key === 'ParentLink') {
          obj._containingFolder = decodeURIComponent(url.parse(cell.Value).pathname);
        }
      }
    });

    if (data.length <= 10) data.push(obj);
  });

  return data;
}

function getSearchPath(options) {
  // Remove trailing slash which breaks the subsite path in the query
  let subsitePath = options.subsite.trim().endsWith('/') ? options.subsite.trim().slice(0, -1) : options.subsite.trim();
  if (subsitePath.startsWith('http://') || subsitePath.startsWith('https://')) {
    return `path:${encodeURI(subsitePath)}`;
  } else if (subsitePath.startsWith('/')) {
    return `path:${encodeURI(options.host + subsitePath)}`;
  } else {
    return `path:${encodeURI(options.host + '/' + subsitePath)}`;
  }
}

function maybeSetClientApplication(options) {
  Logger.trace({ options }, 'maybeSetClientApplication');

  if (clientApplication === null) {
    // The private key is currently retrieved from a text input which does not preserve
    // newlines. For the key to work properly we have to re-add the newlines.
    let keyInput = options.privateKey.trim();

    // Robustly extract and reconstruct the PEM key
    const pemMatch = keyInput.match(/-----BEGIN PRIVATE KEY-----([\s\S]+?)-----END PRIVATE KEY-----/);
    if (!pemMatch) {
      throw new Error('Invalid private key format: missing or malformed BEGIN/END PRIVATE KEY markers.');
    }
    let base64Content = pemMatch[1].replace(/[\r\n\s]/g, '');
    // Validate base64 content
    if (!/^[A-Za-z0-9+/=]+$/.test(base64Content)) {
      throw new Error('Invalid private key format: non-base64 content detected.');
    }
    // Wrap base64 content at 64 characters per line
    const wrappedContent = base64Content.replace(/(.{1,64})/g, '$1\n').trim();
    const key = '-----BEGIN PRIVATE KEY-----\n' + wrappedContent + '\n-----END PRIVATE KEY-----';
    const privateKeyOptions = {
      key: key,
      format: 'pem'
    };

    if (options.privateKeyPassphrase && options.privateKeyPassphrase.length > 0) {
      privateKeyOptions.passphrase = options.privateKeyPassphrase;
    }

    Logger.trace({ privateKeyOptions }, 'Private Key Options');

    const privateKeyObject = crypto.createPrivateKey(privateKeyOptions);
    const privateKey = privateKeyObject.export({
      format: 'pem',
      type: 'pkcs8'
    });

    // The msal-node library uses its own HTTPClient which does not have proper proxy support.
    // As a result, we have to implement our own HTTPClient that wraps the postman-request library
    // to get proper proxy support.
    const customHttpClient = new SharepointHttpAuthClient(requestWithDefaults, Logger);

    const clientConfig = {
      auth: {
        clientId: options.clientId,
        authority: `https://login.microsoftonline.com/${options.tenantId}`,
        clientCertificate: {
          thumbprint: options.publicKeyThumbprint,
          privateKey
        }
      },
      system: {
        loggerOptions: {
          loggerCallback(logLevel, message, containsPii) {
            Logger[msalLogLevelToPolarity(logLevel)]({ logLevel, message, containsPii }, 'MSAL Logger');
          },
          piiLoggingEnabled: config.logging.level === 'trace' ? true : false,
          logLevel: polarityToMsalLogLevel(config.logging.level)
        },
        networkClient: customHttpClient.getClient()
      }
    };

    Logger.trace({ clientConfig }, 'MSAL Client Config');

    clientApplication = new msal.ConfidentialClientApplication(clientConfig);
  }
}

function polarityToMsalLogLevel(polarityLogLevel) {
  switch (polarityLogLevel) {
    case 'trace':
      return msal.LogLevel.Verbose;
    case 'debug':
      return msal.LogLevel.Verbose;
    case 'info':
      return msal.LogLevel.Info;
    case 'warn':
      return msal.LogLevel.Warning;
    case 'error':
      return msal.LogLevel.Error;
    default:
      return msal.LogLevel.Info;
  }
}

function msalLogLevelToPolarity(msalLogLevel) {
  switch (msalLogLevel) {
    case msal.LogLevel.Verbose:
      return 'trace';
    case msal.LogLevel.Info:
      return 'info';
    case msal.LogLevel.Warning:
      return 'warn';
    case msal.LogLevel.Error:
      return 'error';
    default:
      return 'info';
  }
}

async function getToken(options) {
  const clientCredentialRequest = {
    //scopes: ['https://graph.microsoft.com/.default'] //clientCredentialRequestScopes,
    scopes: [`${options.host}/.default`],
    // azureRegion: ro ? ro.region : null, // (optional) specify the region you will deploy your application to here (e.g. "westus2")
    skipCache: true // (optional) this skips the cache and forces MSAL to get a new token from Azure AD
  };

  const response = await clientApplication.acquireTokenByClientCredential(clientCredentialRequest);

  Logger.trace({ response }, 'Acquire Tokens Response');
  return response.accessToken;
}

const getTokenCacheKey = (options) => options.host + options.tenantId + options.clientId;

async function getAuthToken(options) {
  let tokenCacheKey = getTokenCacheKey(options);
  let token = cache.get(tokenCacheKey);

  if (token) {
    Logger.trace('Using cached token for lookup');
    return token;
  }

  maybeSetClientApplication(options);

  let newToken = await getToken(options);
  cache.set(tokenCacheKey, newToken);
  return newToken;
}

function querySharepoint(entity, token, options, callback) {
  let querytext;
  if (options.subsite) {
    querytext = `'${getSearchPath(options)} ${options.directSearch ? `"${entity.value}"` : entity.value}'`;
  } else {
    querytext = options.directSearch ? `'"${entity.value}"'` : `'${entity.value}'`;
  }

  const requestOptions = {
    url: `${options.host}/_api/search/query`,
    qs: {
      querytext,
      RowLimit: 10
    },
    headers: {
      Authorization: 'Bearer ' + token
    }
  };

  Logger.trace({ requestOptions }, 'Sharepoint query request options');

  let totalRetriesLeft = 4;
  const requestSharepoint = () => {
    requestWithDefaults(requestOptions, (err, response, body) => {
      if (err) return callback(err);

      let { statusCode, headers } = response;
      const retryAfter = headers['Retry-After'] || headers['retry-after'];
      if (statusCode === 200) {
        //Logger.trace({ headers }, 'Results of Sharepoint query headers');

        callback(null, body);
      } else if ([429, 500, 503].includes(statusCode) && totalRetriesLeft) {
        totalRetriesLeft--;
        setTimeout(requestSharepoint, retryAfter * 1000 || 1000);
      } else {
        Logger.error({ err, body }, 'Error querying Sharepoint');
        callback(new Error('status code was ' + statusCode));
      }
    });
  };

  requestSharepoint();
}

const parseErrorToReadableJSON = (error) => JSON.parse(JSON.stringify(error, Object.getOwnPropertyNames(error)));

async function doLookup(entities, options, callback) {
  options.subsite = options.subsite.startsWith('/') ? options.subsite.slice(1) : options.subsite;

  Logger.trace({ options }, 'doLookup options');

  _setupRegexBlocklists(options);

  let token;
  try {
    token = await getAuthToken(options);
  } catch (authError) {
    let jsonError = parseErrorToReadableJSON(authError);
    Logger.error({ jsonError }, 'Error getting auth token');
    callback({
      detail: 'Error getting Auth token',
      jsonError
    });
    return;
  }

  const requestQueue = entities.reduce((requestQueue, entity) => {
    if (_isEntityBlocklisted(entity)) return requestQueue;

    // We have to do 1 request per query because we can only AND the query
    // params not OR them
    return requestQueue.concat((done) =>
      querySharepoint(entity, token, options, async (err, body) => {
        if (err) return done(err);

        //Logger.trace({ entity, body }, 'Results of Sharepoint query');

        if (body.PrimaryQueryResult.RelevantResults.RowCount < 1) return done(null, { entity, data: null });

        let details = formatSearchResults(body, options);
        if (details.length < 1) return done(null, { entity, data: null });
        let tags = details.slice(0, 3).map(({ Title, FileType }) => `${Title}.${FileType}`);
        if (details.length > 3) {
          tags.push(`${details.length - 3} more results`);
        }

        done(null, {
          entity,
          data: {
            summary: tags,
            details
          }
        });
      })
    );
  }, []);

  async.parallelLimit(requestQueue, MAX_PARALLEL_LOOKUPS, (err, lookupResults) => {
    if (err) {
      Logger.error('lookup errored', err);

      // errors can sometime have circular structure and this breaks polarity
      callback({ detail: util.inspect(err) });
      return;
    }

    Logger.trace({ lookupResults }, 'Results');

    callback(null, lookupResults);
  });
}

function startup(logger) {
  Logger = logger;
  let defaults = {};

  defaults.json = true;

  requestWithDefaults = request.defaults(defaults);
}

function validateStringOption(errors, options, optionName, errMessage) {
  if (
    typeof options[optionName].value !== 'string' ||
    (typeof options[optionName].value === 'string' && options[optionName].value.length === 0)
  ) {
    errors.push({
      key: optionName,
      message: errMessage
    });
  }
}

function validateOptions(options, callback) {
  let errors = [];

  validateStringOption(errors, options, 'host', 'You must provide a Host option.');
  validateStringOption(errors, options, 'clientId', 'You must provide a Client ID option.');
  validateStringOption(errors, options, 'tenantId', 'You must provide a Tenant ID option.');
  validateStringOption(errors, options, 'publicKeyThumbprint', 'You must provide a Public Key Thumbprint.');
  validateStringOption(errors, options, 'privateKey', 'You must provide Private Key Content.');

  const subsiteStartWithError =
    options.subsite.value && options.subsite.value.startsWith('//')
      ? [
          {
            key: 'subsite',
            message: 'Your subsite must not start with a //'
          }
        ]
      : [];

  callback(null, errors.concat(subsiteStartWithError));
}

module.exports = {
  doLookup,
  startup,
  validateOptions
};
