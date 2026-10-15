const { dispatch } = require('../dist/dispatch');

module.exports = async function (context, req) {
  await dispatch(context, req, 'admin-alerts');
};
