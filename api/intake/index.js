const { handleIntake } = require('../dist/intakeHandler');

module.exports = async function (context, req) {
  await handleIntake(context, req);
};
