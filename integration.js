const async = require("async");
const config = require("./config/config");
const request = require("request");
const util = require("util");
const url = require("url");
const fs = require("fs");
const NodeCache = require("node-cache");
const cache = new NodeCache({
  stdTTL: 60 * 10
});
const MAX_PARALLEL_LOOKUPS = 10;
let Logger;
let requestOptions = {};

let domainBlackList = [];
let previousDomainBlackListAsString = "";
let previousDomainRegexAsString = "";
let previousIpRegexAsString = "";
let domainBlacklistRegex = null;
let ipBlacklistRegex = null;

function _setupRegexBlacklists(options) {
  if (
    options.domainBlacklistRegex !== previousDomainRegexAsString &&
    options.domainBlacklistRegex.length === 0
  ) {
    Logger.debug("Removing Domain Blacklist Regex Filtering");
    previousDomainRegexAsString = "";
    domainBlacklistRegex = null;
  } else {
    if (options.domainBlacklistRegex !== previousDomainRegexAsString) {
      previousDomainRegexAsString = options.domainBlacklistRegex;
      Logger.debug(
        { domainBlacklistRegex: previousDomainRegexAsString },
        "Modifying Domain Blacklist Regex"
      );
      domainBlacklistRegex = new RegExp(options.domainBlacklistRegex, "i");
    }
  }

  if (
    options.blacklist !== previousDomainBlackListAsString &&
    options.blacklist.length === 0
  ) {
    Logger.debug("Removing Domain Blacklist Filtering");
    previousDomainBlackListAsString = "";
    domainBlackList = null;
  } else {
    if (options.blacklist !== previousDomainBlackListAsString) {
      previousDomainBlackListAsString = options.blacklist;
      Logger.debug(
        { domainBlacklist: previousDomainBlackListAsString },
        "Modifying Domain Blacklist Regex"
      );
      domainBlackList = options.blacklist.split(",").map((item) => item.trim());
    }
  }

  if (
    options.ipBlacklistRegex !== previousIpRegexAsString &&
    options.ipBlacklistRegex.length === 0
  ) {
    Logger.debug("Removing IP Blacklist Regex Filtering");
    previousIpRegexAsString = "";
    ipBlacklistRegex = null;
  } else {
    if (options.ipBlacklistRegex !== previousIpRegexAsString) {
      previousIpRegexAsString = options.ipBlacklistRegex;
      Logger.debug(
        { ipBlacklistRegex: previousIpRegexAsString },
        "Modifying IP Blacklist Regex"
      );
      ipBlacklistRegex = new RegExp(options.ipBlacklistRegex, "i");
    }
  }
}

function _isEntityBlacklisted(entityObj, options) {
  if (domainBlackList.indexOf(entityObj.value) >= 0) {
    return true;
  }

  if (entityObj.isIPv4 && !entityObj.isPrivateIP) {
    if (ipBlacklistRegex !== null) {
      if (ipBlacklistRegex.test(entityObj.value)) {
        Logger.debug({ ip: entityObj.value }, "Blocked BlackListed IP Lookup");
        return true;
      }
    }
  }

  if (entityObj.isDomain) {
    if (domainBlacklistRegex !== null) {
      if (domainBlacklistRegex.test(entityObj.value)) {
        Logger.debug({ domain: entityObj.value }, "Blocked BlackListed Domain Lookup");
        return true;
      }
    }
  }

  return false;
}

function formatSearchResults(searchResults) {
  let data = [];

  searchResults.PrimaryQueryResult.RelevantResults.Table.Rows.forEach((row) => {
    let obj = {};
    row.Cells.forEach((cell) => {
      obj[cell.Key] = cell.Value;
    });

    data.push(obj);
  });

  return data;
}

function getRequestOptions() {
  return JSON.parse(JSON.stringify(requestOptions));
}

const getTokenCacheKey = (options) =>
  options.host +
  options.authHost +
  options.tenantId +
  options.clientId +
  options.clientSecret;

function getAuthToken(options, callback) {
  let tokenCacheKey = getTokenCacheKey(options);
  let token = cache.get(tokenCacheKey);

  if (token) {
    callback(null, token);
    return;
  }

  let hostUrl = url.parse(options.host);

  request({
    url: `${options.authHost}/${options.tenantId}/tokens/OAuth/2`,
    formData: {
      grant_type: "client_credentials",
      client_id: `${options.clientId}@${options.tenantId}`,
      client_secret: options.clientSecret,
      resource: `00000003-0000-0ff1-ce00-000000000000/${hostUrl.host}@${options.tenantId}`
    },
    json: true,
    method: "POST"
  },(err, resp, body) => {
    if (err) return callback(err);

    if (resp.statusCode !== 200) {
      callback({ err: new Error("status code was not 200"), body: body });
      return;
    }

    cache.set(tokenCacheKey, body.access_token);

    callback(null, body.access_token);
  });
}

function querySharepoint(entity, token, options, callback) {
  let requestOptions = getRequestOptions();

  if (options.subsite) {
    requestOptions.qs = {
      querytext: `'path:${options.host}/${options.subsite} ${entity.value}'`
    };
  } else {
    requestOptions.qs = {
      querytext: `'${entity.value}'`
    };
  }
  requestOptions.url = `${options.host}/_api/search/query`;
  requestOptions.headers = {
    Authorization: "Bearer " + token
  };

  const requestSharepoint = () => {
    const sharepointRetryDateTime = cache.getTtl("sharepointIsThrottled");

    if (sharepointRetryDateTime) {
      const waitTime = (sharepointRetryDateTime - Date.now());
      return setTimeout(requestSharepoint, waitTime);
    }

    request(requestOptions, (err, { statusCode, headers }, body) => {
      if (err) return callback(err);
      const retryAfter = headers["Retry-After"];
      if (statusCode === 200) {
        Logger.trace({ headers }, "Results of Sharepoint qeury headers");

        callback(null, body);
      } else if ((statusCode === 429 || statusCode === 503) && retryAfter) {
        cache.set("sharepointIsThrottled", true, retryAfter);

        setTimeout(requestSharepoint, retryAfter * 1000);
      } else {
        callback(new Error("status code was " + statusCode));
      }
    });
  }

  requestSharepoint();
}

function doLookup(entities, options, callback) {
  Logger.trace("starting lookup");

  Logger.trace("options are", options);

  _setupRegexBlacklists(options);

  getAuthToken(options, (err, token) => {
    if (err) {
      Logger.error("get token errored", err);
      callback({ err: err });
      return;
    }

    const requestQueue = entities.reduce((requestQueue, entity) => {
      if (_isEntityBlacklisted(entity)) return requestQueue;
      
      // We have to do 1 request per query because we can only AND the query
      // params not OR them
      return requestQueue.concat((done) =>
        querySharepoint(entity, token, options, (err, body) => {
          if (err) return done(err);

          Logger.trace({ entity, body }, "Results of Sharepoint qeury");

          if (body.PrimaryQueryResult.RelevantResults.RowCount < 1)
            return done(null, { entity, data: null });

          let details = formatSearchResults(body);
          let tags = details.map(({ Title, FileType }) => `${Title}.${FileType}`);

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
        Logger.error("lookup errored", err);

        // errors can sometime have circular structure and this breaks polarity
        callback({ detail: util.inspect(err) });
        return;
      }

      Logger.debug({ lookupResults }, "Results");

      callback(null, lookupResults);
    });
  });
}

function startup(logger) {
  Logger = logger;

  if (typeof config.request.cert === "string" && config.request.cert.length > 0) {
    requestOptions.cert = fs.readFileSync(config.request.cert);
  }

  if (typeof config.request.key === "string" && config.request.key.length > 0) {
    requestOptions.key = fs.readFileSync(config.request.key);
  }

  if (
    typeof config.request.passphrase === "string" &&
    config.request.passphrase.length > 0
  ) {
    requestOptions.passphrase = config.request.passphrase;
  }

  if (typeof config.request.ca === "string" && config.request.ca.length > 0) {
    requestOptions.ca = fs.readFileSync(config.request.ca);
  }

  if (typeof config.request.proxy === "string" && config.request.proxy.length > 0) {
    requestOptions.proxy = config.request.proxy;
  }

  if (typeof config.request.rejectUnauthorized === "boolean") {
    requestOptions.rejectUnauthorized = config.request.rejectUnauthorized;
  }

  requestOptions.json = true;
}

function validateStringOption(errors, options, optionName, errMessage) {
  if (
    typeof options[optionName].value !== "string" ||
    (typeof options[optionName].value === "string" &&
      options[optionName].value.length === 0)
  ) {
    errors.push({
      key: optionName,
      message: errMessage
    });
  }
}

function validateOptions(options, callback) {
  let errors = [];

  validateStringOption(errors, options, "host", "You must provide a Host option.");
  validateStringOption(
    errors,
    options,
    "authHost",
    "You must provide an Authentication Host option."
  );
  validateStringOption(
    errors,
    options,
    "clientId",
    "You must provide a Client ID option."
  );
  validateStringOption(
    errors,
    options,
    "clientSecret",
    "You must provide a Client Secret option."
  );
  validateStringOption(
    errors,
    options,
    "tenantId",
    "You must provide a Tenant ID option."
  );

  callback(null, errors);
}

module.exports = {
  doLookup: doLookup,
  startup: startup,
  validateOptions: validateOptions
};
