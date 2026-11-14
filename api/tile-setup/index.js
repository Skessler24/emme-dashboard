const { handleTileSetup } = require('../dist/tileSetup');

module.exports = async function (context, req) {
  await handleTileSetup(context, req);
};
