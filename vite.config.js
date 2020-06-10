module.exports = {
  proxy: {
    '/api': {
      target: 'http://127.0.0.1:8088',
      changeOrigin: true,
    }
  }
};
