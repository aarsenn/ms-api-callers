const axios = require('axios');
const { v4: uuidv4 } = require('uuid');

const APP_ID_TYPES = {
  INHERIT_FROM_SESSION: 'inherit_from_session'
};

const APP_ID_SESSION_KEY = '__last_ms_teams_app_id';

const CONFIG = {
  eventsAdapterName: 'MS Teams Events Adapter',
  graphAuthRequestEndpoint: 'https://login.microsoftonline.com/common/adminconsent',
  graphAuthRequestScope: 'https://graph.microsoft.com/.default',
  graphGetTokenEndpoint: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token',
  graphApiEndpoint: 'https://graph.microsoft.com/v1.0/',
  botGetTokenEndpoint: 'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token',
  botGetTokenScope: 'https://api.botframework.com/.default',
  botDefaultApiEndpoint: 'https://smba.trafficmanager.net/amer/',
  tableName: '__ms_teams_auth__',
  listKey: '__list',
  installerBotName: 'installer:teams'
};

const createAppIdProvider = (persist, msAuthAppSelectMethod) => ({
  get: async appId => {
    if (msAuthAppSelectMethod !== 'inherited') return appId;

    appId = persist && persist[APP_ID_SESSION_KEY];

    if (!appId) {
      throw new Error(`Can't inherit appId from session.`);
    }

    return appId;
  },
  set: async appId => {
    if (!appId) {
      throw new Error(`Can't set ms teams app id in session.`)
    }
    persist = persist || {};
    return persist[APP_ID_SESSION_KEY] = appId;
  }
});

async function initApiCallers({ appId, msAuthAppSelectMethod, storage, context, persist, appUpdatedCallback }) {
  try {
    const config = CONFIG;
    if (!config) throw new Error(`Can't get config on API callers init`);

    const appIdProvider = createAppIdProvider(persist, msAuthAppSelectMethod);

    let app = await storage.get(
      config.tableName,
      await appIdProvider.get(appId)
    );

    if (!app) throw new Error(`Can't get app on API callers init`);

    await appIdProvider.set(app.id);

    const authProviderUrl = context.helpers.gatewayUrl('msteams/oauth', context.helpers.providersAccountId);

    const isFunction = f => typeof f === 'function';

    const emitAppUpdated = app => {
      if (!isFunction(appUpdatedCallback)) return;
      appUpdatedCallback(app);
    };

    const hideToken = (requestOptions, replacer = 'TOKEN IS HIDDEN') => ({
      ...requestOptions,
      headers: {
        ...requestOptions.headers,
        Authorization: requestOptions.headers.Authorization ? replacer : requestOptions.headers.Authorization
      }
    });

    const checkTokens = () => {
      if (!app) return Promise.reject(new Error('No app providet in check tokens'));
      if (app.tokensExpired >= Date.now()) return Promise.resolve(app);

      const requestOptions = {
        url: authProviderUrl,
        method: 'post',
        data: { app },
        headers: { 'Authorization': `FLOW ${context.config.flowToken}` }
      };

      context.log.debug('Start MS Teams tokens refresh', { app, callingFlow: authProviderUrl });

      return axios(requestOptions).then(resp => {
        if (!resp || !resp.data) return null;
        app = resp.data;
        emitAppUpdated(app);
        context.log.debug('Tokens updated', { app });
        return resp.data;
      });
    };

    const replVers = (str, val) => {
      return str.replace(/v\d\.\d/, val);
    };

    const callGraph = opts => {
      const url = url => opts.version ? replVers(url, opts.version) : url;
      const requestId = uuidv4();
      return checkTokens()
        .then(app => {
          const headers = { 'Authorization': `Bearer ${app.graphToken}` };
          const graphOpts = {
            headers: Object.assign(headers, opts.headers || {}),
            url: url(config.graphApiEndpoint) + opts.url,
          };
          const httpOpts = Object.assign(opts, graphOpts);

          context.log.debug(`MS Graph request - ${requestId}`, hideToken(httpOpts));

          return axios(httpOpts);
        })
        .then(resp => opts.meta ? resp.data : (resp.data.value || resp.data))
        .then(value => {
          context.log.debug(`MS Graph response - ${requestId}`, value);
          return value;
        })
        .catch(err => {
          const requestError = () => ({ response: err.response.data, status: err.response.status });
          const error = () => ({ message: err.message || err });
          context.log.error(`Error on MS Graph request - ${requestId}`, err.response ? requestError() : error());
          throw err;
        });
    };

    const callBotframework = opts => {
      const url = url => opts.version ? replVers(url, opts.version) : url;
      const requestId = uuidv4();

      return checkTokens()
        .then(app => {
          const headers = { 'Authorization': `Bearer ${app.botToken}` };
          const graphOpts = {
            headers: Object.assign(headers, opts.headers || {}),
            url: url(app.botServiceUrl || config.botDefaultApiEndpoint) + opts.url,
          };
          const httpOpts = Object.assign(opts, graphOpts);
          context.log.debug(`MS Botframework request - ${requestId}`, hideToken(httpOpts));
          return axios(httpOpts);
        })
        .then(resp => resp.data)
        .then(value => {
          context.log.debug(`MS Botframework response - ${requestId}`, value);
          return value;
        })
        .catch(err => {
          const requestError = () => ({ response: err.response.data, status: err.response.status });
          const error = () => ({ message: err.message || err });
          context.log.error(`Error on MS Botframework request - ${requestId}`, err.response ? requestError() : error());
          throw err;
        });
    };

    return {
      callGraph: opts => callGraph(isFunction(opts) ? opts(app) : opts),
      callBotframework: opts => callBotframework(isFunction(opts) ? opts(app) : opts),
      app
    }
  } catch (error) {
    throw error;
  }
};

exports.initApiCallers = initApiCallers;