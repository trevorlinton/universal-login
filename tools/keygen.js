'use strict';

const fs = require('fs');
const path = require('path');
const { createKeyStore } = require('oidc-provider');

const keystore = createKeyStore();

Promise.all([
  keystore.generate('RSA', 2048),
  keystore.generate('EC', 'P-256'),
  keystore.generate('EC', 'P-384'),
  keystore.generate('EC', 'P-521'),
]).then(() => {
  console.log((new Buffer(JSON.stringify(keystore.toJSON(true)))).toString('base64'));
});