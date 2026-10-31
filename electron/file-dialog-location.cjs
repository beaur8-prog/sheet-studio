'use strict';
const path = require('node:path');

// AppKit restores its own bookmark before Electron can apply defaultPath.
// Use a plain path for this app's location to avoid synchronous cloud-bookmark
// resolution during panel creation. Never change the shared panel preferences.
function prepareFileDialogLocation(systemPreferences, directory, platform = process.platform) {
  if (platform !== 'darwin' || typeof directory !== 'string' || !path.isAbsolute(directory)) return;
  systemPreferences.setUserDefault('NSOSPLastRootDirectory', 'string', directory);
}
module.exports = { prepareFileDialogLocation };
