'use strict';
const { fileAssociations } = require('../electron/file-open-requests.cjs');

const classes = 'Software\\Classes\\';
const application = classes + 'Applications\\Sheet Studio.exe';
const capabilities = 'Software\\Sheet Studio\\Capabilities';
const executable = String.raw`$INSTDIR\Sheet Studio.exe`;
const command = String.raw`$\"$INSTDIR\Sheet Studio.exe$\" $\"%1$\"`;
const icon = String.raw`$\"$INSTDIR\Sheet Studio.exe$\",0`;
const description = 'Spreadsheet editor for XLSX, XLSM, CSV and TSV files';
const endIf = '${EndIf}';
const write = (key, name, value) => `  WriteRegStr SHELL_CONTEXT "${key}" "${name}" '${value}'`;
const remove = (key, name) => `  DeleteRegValue SHELL_CONTEXT "${key}" "${name}"`;
// The bundled NSIS can treat /ifempty as "no subkeys", even when values
// remain. Explicitly enumerate BOTH before considering a key for deletion.
const prune = keys => keys.map(key => `
  ClearErrors
  EnumRegValue $0 SHELL_CONTEXT "${key}" 0
  \${If} \${Errors}
    ClearErrors
    EnumRegKey $0 SHELL_CONTEXT "${key}" 0
    \${If} \${Errors}
      DeleteRegKey /ifempty SHELL_CONTEXT "${key}"
    ${endIf}
  ${endIf}
  ClearErrors`).join('\n');

// Do not remove a value another installer changed, or its added verbs/data.
const removeMatching = (key, name, value) => `
  ReadRegStr $0 SHELL_CONTEXT "${key}" "${name}"
  \${If} $0 == '${value}'
${remove(key, name)}
  ${endIf}`;

function windowsFileAssociations() {
  const install = [], uninstall = [];
  for (const { ext, name } of fileAssociations) {
    const progId = `SheetStudio.${ext}`, key = classes + progId;
    const legacy = classes + name;
    install.push([
      write(key, '', name), write(key + '\\DefaultIcon', '', icon),
      write(key + '\\shell\\open\\command', '', command),
      write(key + '\\Application', 'ApplicationName', 'Sheet Studio'),
      `  WriteRegNone SHELL_CONTEXT "${classes}.${ext}\\OpenWithProgids" "${progId}"`,
      `  WriteRegNone SHELL_CONTEXT "${application}\\SupportedTypes" ".${ext}"`,
      write(capabilities + '\\FileAssociations', '.' + ext, progId),
      // Previous versions used generic type names. Add only missing commands
      // and icons, preserving other applications' existing data.
      `  ReadRegStr $0 SHELL_CONTEXT "${legacy}\\shell\\open\\command" ""`,
      '${If} $0 == ""',
      write(legacy, 'SheetStudioLegacyOwner', executable),
      write(legacy + '\\shell\\open\\command', '', command),
      `  ReadRegStr $0 SHELL_CONTEXT "${legacy}\\DefaultIcon" ""`,
      '${If} $0 == ""',
      write(legacy + '\\DefaultIcon', '', icon),
      write(legacy, 'SheetStudioLegacyIconOwner', executable),
      endIf, endIf,
    ].join('\n'));
    uninstall.push([
      remove(classes + '.' + ext + '\\OpenWithProgids', progId),
      removeMatching(key, '', name), removeMatching(key + '\\DefaultIcon', '', icon),
      removeMatching(key + '\\shell\\open\\command', '', command),
      removeMatching(key + '\\Application', 'ApplicationName', 'Sheet Studio'),
      prune([key + '\\Application', key + '\\DefaultIcon', key + '\\shell\\open\\command', key + '\\shell\\open', key + '\\shell', key]),
      remove(application + '\\SupportedTypes', '.' + ext),
      removeMatching(capabilities + '\\FileAssociations', '.' + ext, progId),
      `  ReadRegStr $0 SHELL_CONTEXT "${legacy}" "SheetStudioLegacyOwner"`,
      `\${If} $0 == '${executable}'`,
      removeMatching(legacy + '\\shell\\open\\command', '', command),
      removeMatching(legacy, 'SheetStudioLegacyOwner', executable), endIf,
      `  ReadRegStr $0 SHELL_CONTEXT "${legacy}" "SheetStudioLegacyIconOwner"`,
      `\${If} $0 == '${executable}'`,
      removeMatching(legacy + '\\DefaultIcon', '', icon),
      removeMatching(legacy, 'SheetStudioLegacyIconOwner', executable), endIf,
      // Only empty keys may be pruned, including our own ProgID containers.
      prune([legacy + '\\shell\\open\\command', legacy + '\\shell\\open', legacy + '\\shell', legacy + '\\DefaultIcon', legacy]),
    ].join('\n'));
  }
  return `; Generated from the supported spreadsheet types at package time.
!macro customInstall
${write(application, 'FriendlyAppName', 'Sheet Studio')}
${write(application + '\\DefaultIcon', '', icon)}
${write(application + '\\shell\\open\\command', '', command)}
${write(capabilities, 'ApplicationName', 'Sheet Studio')}
${write(capabilities, 'ApplicationDescription', description)}
${write(capabilities, 'ApplicationIcon', icon)}
${write('Software\\RegisteredApplications', 'Sheet Studio', capabilities)}
${install.join('\n')}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
${uninstall.join('\n')}
${removeMatching(application, 'FriendlyAppName', 'Sheet Studio')}
${removeMatching(application + '\\DefaultIcon', '', icon)}
${removeMatching(application + '\\shell\\open\\command', '', command)}
${prune([application + '\\DefaultIcon', application + '\\shell\\open\\command', application + '\\shell\\open', application + '\\shell', application + '\\SupportedTypes', application])}
${removeMatching('Software\\RegisteredApplications', 'Sheet Studio', capabilities)}
${removeMatching(capabilities, 'ApplicationName', 'Sheet Studio')}
${removeMatching(capabilities, 'ApplicationDescription', description)}
${removeMatching(capabilities, 'ApplicationIcon', icon)}
${prune([capabilities + '\\FileAssociations', capabilities])}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
`;
}

module.exports = { windowsFileAssociations };
