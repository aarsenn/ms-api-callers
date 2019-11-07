const axios = require('axios');
// const qs = require('qs');

const CONFIG_TABLE = '__ms_teams_auth_config__';
const CONFIG_KEY = '__config';

const checkTokens = app => {
  if (!app) return Promise.reject(new Error('No app providet in check tokens'));
  if (app.tokensExpired >= Date.now()) return Promise.resolve(app);
  return axios({
    url: app.adapterUrl,
    method: 'post',
    data: { app } 
  }).then(resp => {
    if (!resp.data) return null;
    return resp.data;
  });
}

const callGraph = (opts, app, config, context) => {
  return checkTokens(app)
    .then(app => {
      const headers = { 'Authorization': `Bearer ${app.graphToken}` };
      const graphOpts = {
        headers: Object.assign(headers, opts.headers || {}),
        url: config.graphApiEndpoint + opts.url,
      };
      const httpOpts = Object.assign(opts, graphOpts);
      context.log.info('MS Graph http request', { request: httpOpts });
      return axios(httpOpts);
    })
    .then(resp => resp.data.value || resp.data)
    .catch(err => {
      context.log.error('Error on http request', { message: err.message, stack: err.stack });
      throw err;
    });
};

const callBotframework = (opts, app, config, context) => {
  return checkTokens(app)
    .then(app => {
      const headers = { 'Authorization': `Bearer ${app.botToken}` };
      const graphOpts = {
        headers: Object.assign(headers, opts.headers || {}),
        url: (app.botServiceUrl || config.botDefaultApiEndpoint) + opts.url,
      };
      const httpOpts = Object.assign(opts, graphOpts);
      context.log.info('MS Botframework http request', { request: httpOpts });
      return axios(httpOpts);
    })
    .then(resp => resp.data)
    .catch(err => {
      context.log.error('Error on http request', { message: err.message, stack: err.stack });
      throw err;
    });
};

async function initApiCallers(appId, storage) {
  try {
    const config = await storage.get(CONFIG_TABLE, CONFIG_KEY);
    if (!config) throw new Error(`Can't get config on API callers init`);

    const app = await storage.get(config.tableName, appId);
    if (!app) throw new Error(`Can't get app on API callers init`);

    const isFunc = f => typeof f === 'function';

    return {
      callGraph: opts => callGraph(isFunc(opts) ? opts(app) : opts, app, config, this),
      callBotframework: opts => callBotframework(isFunc(opts) ? opts(app) : opts, app, config, this),
      app
    }
  } catch (error) {
    throw error;
  }
};

exports.initApiCallers = initApiCallers;