const { forward } = require('./_bridge');

module.exports = async (req, res) => {
  return forward('/health', req, res);
};
