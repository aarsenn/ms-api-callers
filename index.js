const axios = require('axios');
// const qs = require('qs');

const CONFIG_TABLE = '__ms_teams_auth_config__';
const CONFIG_KEY = '__config';

function checkTokens(app) {
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

const callGraph = (opts, app, config) => {
  return checkTokens(app)
    .then(app => {
      const headers = { 'Authorization': `Bearer ${app.graphToken}` };
      const graphOpts = {
        headers: _.assign(headers, opts.headers || {}),
        url: config.graphApiEndpoint + opts.url,
      };
      const httpOpts = _.assign(opts, graphOpts);
      return axios(httpOpts);
    });
};

const callBotframework = (opts, app, config) => {
  return checkTokens(app)
    .then(app => {
      const headers = { 'Authorization': `Bearer ${app.botToken}` };
      const graphOpts = {
        headers: _.assign(headers, opts.headers || {}),
        url: (app.botServiceUrl || config.botDefaultApiEndpoint) + opts.url,
      };
      const httpOpts = _.assign(opts, graphOpts);
      return axios(httpOpts);
    });
};

async function initApiCallers(appId, storage) {
  try {
    const config = await storage.get(CONFIG_TABLE, CONFIG_KEY);
    this.log.warn('CONFIG', { config });
    if (!config) throw new Error(`Can't get config on API callers init`);

    const app = await storage.get(config.tableName, appId);
    this.log.warn('APP', { app });
    if (!app) throw new Error(`Can't get app on API callers init`);

    return {
      callGraph: opts => callGraph(opts, app, config),
      callBotframework: opts => callBotframework(opts, app, config)
    }
  } catch (error) {
    throw error;
  }
};

exports.initApiCallers = initApiCallers;