// Shim re-export so tsserver can resolve `require('office-scripts-plugin')`.
//
// Two layouts we need to support:
//
//  a) Dev: npm's `file:` dep creates a symlink
//     <ext>/node_modules/office-scripts-plugin -> <ext>/plugin-package
//     Node realpath's through the symlink by default, so __dirname here is
//     <ext>/plugin-package and the bundled plugin lives at ../dist/plugin.js.
//
//  b) Installed VSIX: vsce copies plugin-package's contents into the archive
//     verbatim, so __dirname is <ext>/node_modules/office-scripts-plugin with
//     no symlink to follow. The bundled plugin lives at ../../dist/plugin.js.
//
// Detect by asking: are we inside a real node_modules/ path? If so, walk up
// two levels; otherwise walk up one.
const path = require('path');

const insideNodeModules = __dirname
    .split(path.sep)
    .includes('node_modules');

const pluginPath = insideNodeModules
    ? path.resolve(__dirname, '..', '..', 'dist', 'plugin.js')
    : path.resolve(__dirname, '..', 'dist', 'plugin.js');

module.exports = require(pluginPath);
