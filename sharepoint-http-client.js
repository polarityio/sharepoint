/*
 * Copyright (c) 2023, Polarity.io, Inc.
 */

/**
 * Implements the msal-node `networkClient` interface using `postman-request` as the underlying HTTP client so that
 * proxy support is available.
 * See: https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/2600#issuecomment-831299149
 *
 * The msal-node library originally used gaxios which does not have proper proxy support.  To work around this
 * msal-node implemented their own `networkClient` interface which should support proxies but it does not always
 * work.
 */
class SharepointHttpAuthClient {
  constructor(requestWithDefaults, Logger) {
    this.requestWithDefaults = requestWithDefaults;
    this.logger = Logger;
  }

  getClient() {
    return {
      sendGetRequestAsync: this.sendGetRequestAsync.bind(this),
      sendPostRequestAsync: this.sendPostRequestAsync.bind(this)
    };
  }

  async sendGetRequestAsync(uri, options) {
    this.logger.trace({ uri, options }, 'sendGetRequestAsync');
    return this.sendAsyncRequest(uri, 'GET', options);
  }

  async sendPostRequestAsync(uri, options) {
    this.logger.trace({ uri, options }, 'sendPostRequestAsync');
    return this.sendAsyncRequest(uri, 'POST', options);
  }

  async sendAsyncRequest(uri, method, options) {
    return new Promise((resolve, reject) => {
      const requestOptions = {
        uri,
        method,
        json: true,
        ...options
      };
      this.requestWithDefaults(requestOptions, (err, response, body) => {
        if (err) {
          return reject(err);
        }

        resolve({
          headers: response.headers,
          body,
          status: response.status
        });
      });
    });
  }
}

module.exports = SharepointHttpAuthClient;
