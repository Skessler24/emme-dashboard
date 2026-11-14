const { handleBillingIntake } = require('../dist/billingIntake');

module.exports = async function (context, req) {
  await handleBillingIntake(context, req);
};
