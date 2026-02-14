const { forward } = require('./_bridge');

module.exports = async (req, res) => {
  return forward('/format', req, res);
};
