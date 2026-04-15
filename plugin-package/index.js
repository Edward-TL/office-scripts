// Shim re-export so tsserver can resolve `require('office-scripts-plugin')`.
// Node realpath's through the symlink created by `npm install` (file: dep),
// so this relative path resolves to <extension-root>/dist/plugin.js.
module.exports = require('../dist/plugin.js');
