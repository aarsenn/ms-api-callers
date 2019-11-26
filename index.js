const axios = require('axios');

const CONFIG_TABLE = '__ms_teams_auth_config__';
const CONFIG_KEY = '__config';

async function initApiCallers({ appId, storage, context, adaptersLinks, appUpdatedCallback }) {
  try {
    const config = await storage.get(CONFIG_TABLE, CONFIG_KEY);
    if (!config) throw new Error(`Can't get config on API callers init`);

    let app = await storage.get(config.tableName, appId);
    if (!app) throw new Error(`Can't get app on API callers init`);

    const authAdapterUrl = adaptersLinks.find(adapter => adapter.label === config.authAdapterName);
    if (!authAdapterUrl) throw new Error(`Can't get auth adapter url`);

    const isFunc = f => typeof f === 'function';

    const updateAppRecord = data => {
      const updated = Object.assign({}, app, data);
      return storage.set(config.tableName, app.id, updated)
        .then(() => updated);
    };

    const checkTokens = () => {
      if (!app) return Promise.reject(new Error('No app providet in check tokens'));
      if (app.tokensExpired >= Date.now()) return Promise.resolve(app);

      context.log.warn('Update tokens start', { app });

      const isCommonApp = !app.clientSecret;
      const url = isCommonApp ? `${context.authProviderUrl}?type=tokens` : authAdapterUrl;

      return axios({
        url,
        method: 'post',
        data: { app },
        headers: { 'Authorization': `FLOW ${context.config.flowToken}` }
      }).then(resp => {
        if (!resp.data) return null;

        if (isCommonApp) return updateAppRecord(resp.data)
          .then(updatedApp => {
            app = updatedApp;
            appUpdatedCallback(updatedApp);
            context.log.warn('Update tokens finish', { app });
            return updatedApp;
          });
        
        app = resp.data;
        appUpdatedCallback(app);
        context.log.warn('Update tokens finish', { app });
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
        .then(resp => resp.data.value || resp.data)
        .catch(err => {
          context.log.error('Error on http request', { response: err.response.data, status: err.response.status });
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
          context.log.error('Error on http request', { response: err.response.data, status: err.response.status });
          throw err;
        });
    };

    return {
      callGraph: opts => callGraph(isFunc(opts) ? opts(app) : opts),
      callBotframework: opts => callBotframework(isFunc(opts) ? opts(app) : opts),
      app
    }
  } catch (error) {
    throw error;
  }
};

exports.initApiCallers = initApiCallers;