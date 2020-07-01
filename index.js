const axios = require('axios');

const APP_ID_TYPES = {
  INHERIT_FROM_SESSION: 'inherit_from_session'
};
const APP_ID_SESSION_KEY = '__last_ms_teams_app_id';

const CONFIG_TABLE = '__ms_teams_auth_config__';
const CONFIG_KEY = '__config';

const createAppIdProvider = context => ({
  get: async appId => {
    if ((appId !== APP_ID_TYPES.INHERIT_FROM_SESSION)) return appId;
    appId = await context.get(APP_ID_SESSION_KEY);
    if (!appId) {
      throw new Error(`Can't inherit appId from session.`);
    }
    return appId;
  },
  set: async appId => {
    if (!appId) {
      throw new Error(`Can't set ms teams app id in session.`)
    }
    return context.set(APP_ID_SESSION_KEY, appId);
  }
});

async function initApiCallers({ appId, storage, context, appUpdatedCallback }) {
  try {
    const config = await storage.get(CONFIG_TABLE, CONFIG_KEY);
    if (!config) throw new Error(`Can't get config on API callers init`);

    const appIdProvider = createAppIdProvider(context);

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

    const checkTokens = () => {
      if (!app) return Promise.reject(new Error('No app providet in check tokens'));
      if (app.tokensExpired >= Date.now()) return Promise.resolve(app);

      const requestOptions = {
        url: authProviderUrl,
        method: 'post',
        data: { app },
        headers: { 'Authorization': `FLOW ${context.config.flowToken}` }
      };

      context.log.info('Start tokens update', { app, callingFlow: authProviderUrl });

      return axios(requestOptions).then(resp => {
        if (!resp || !resp.data) return null;
        app = resp.data;
        emitAppUpdated(app);
        context.log.info('Tokens updated', { app });
        return resp.data;
      });
    };

    const replVers = (str, val) => {
      return str.replace(/v\d\.\d/, val);
    };

    const callGraph = opts => {
      const url = url => opts.version ? replVers(url, opts.version) : url;
      return checkTokens()
        .then(app => {
          const headers = { 'Authorization': `Bearer ${app.graphToken}` };
          const graphOpts = {
            headers: Object.assign(headers, opts.headers || {}),
            url: url(config.graphApiEndpoint) + opts.url,
          };
          const httpOpts = Object.assign(opts, graphOpts);
          context.log.info('MS Graph http request', { request: httpOpts });
          return axios(httpOpts);
        })
        .then(resp => opts.meta ? resp.data : (resp.data.value || resp.data))
        .catch(err => {
          const requestError = () => ({ response: err.response.data, status: err.response.status });
          const error = () => ({ message: err.message || err });
          context.log.error('Error on http request', err.response ? requestError() : error());
          throw err;
        });
    };

    const callBotframework = opts => {
      const url = url => opts.version ? replVers(url, opts.version) : url;
      return checkTokens()
        .then(app => {
          const headers = { 'Authorization': `Bearer ${app.botToken}` };
          const graphOpts = {
            headers: Object.assign(headers, opts.headers || {}),
            url: url(app.botServiceUrl || config.botDefaultApiEndpoint) + opts.url,
          };
          const httpOpts = Object.assign(opts, graphOpts);
          context.log.info('MS Botframework http request', { request: httpOpts });
          return axios(httpOpts);
        })
        .then(resp => resp.data)
        .catch(err => {
          const requestError = () => ({ response: err.response.data, status: err.response.status });
          const error = () => ({ message: err.message || err });
          context.log.error('Error on http request', err.response ? requestError() : error());
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