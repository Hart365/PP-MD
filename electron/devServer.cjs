function getDevServerUrl() {
  return process.env.PPMD_DEV_SERVER_URL || 'http://127.0.0.1:5173';
}

module.exports = {
  getDevServerUrl,
};
