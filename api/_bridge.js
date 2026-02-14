const { server } = require('../kimi-analyze-server');

function forward(path, req, res) {
  req.url = path;
  server.emit('request', req, res);
}

module.exports = { forward };
