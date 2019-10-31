const CONFIG_TABLE = '__ms_teams_auth_config__';
const CONFIG_KEY = '__config';

const httpCall = () => {

};

const callGraph = (opts, app, config) => {
  return Promise.resolve('Graph ok')
};

const callBotframework = (opts, app, config) => {
  return Promise.resolve('Bot ok')
};

export default async function initApiCallers(appId, storage) {
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