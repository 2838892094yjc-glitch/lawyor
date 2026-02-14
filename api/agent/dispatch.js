const { forward } = require('../_bridge');

module.exports = async (req, res) => {
  return forward('/agent/dispatch', req, res);
};
