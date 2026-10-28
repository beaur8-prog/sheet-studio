'use strict';
const { fileAssociations } = require('../electron/file-open-requests.cjs');

// Register candidates for Open With and Default apps. Do not write extension
// defaults or Explorer UserChoice: only the user chooses the preferred app.
function windowsFileAssociations() {
  const install = [], uninstall = [];
  for (const { ext, name } of fileAssociations) {
    const progId = `SheetStudio.${ext}`;
    install.push(String.raw`
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}" "" "${name}"
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}\DefaultIcon" "" '$\"$INSTDIR\Sheet Studio.exe$\",0'
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}\shell\open\command" "" '$\"$INSTDIR\Sheet Studio.exe$\" $\"%1$\"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}\Application" "ApplicationName" "Sheet Studio"
  WriteRegNone SHELL_CONTEXT "Software\Classes\.${ext}\OpenWithProgids" "${progId}"
  WriteRegNone SHELL_CONTEXT "Software\Classes\Applications\Sheet Studio.exe\SupportedTypes" ".${ext}"
  WriteRegStr SHELL_CONTEXT "Software\Sheet Studio\Capabilities\FileAssociations" ".${ext}" "${progId}"`);
    uninstall.push(String.raw`
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.${ext}\OpenWithProgids" "${progId}"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}"`);
  }
  return String.raw`; Generated from the supported spreadsheet types at package time.
!macro customInstall
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\Sheet Studio.exe" "FriendlyAppName" "Sheet Studio"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\Sheet Studio.exe\DefaultIcon" "" '$\"$INSTDIR\Sheet Studio.exe$\",0'
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\Sheet Studio.exe\shell\open\command" "" '$\"$INSTDIR\Sheet Studio.exe$\" $\"%1$\"'
  WriteRegStr SHELL_CONTEXT "Software\Sheet Studio\Capabilities" "ApplicationName" "Sheet Studio"
  WriteRegStr SHELL_CONTEXT "Software\Sheet Studio\Capabilities" "ApplicationDescription" "Spreadsheet editor for XLSX, XLSM, CSV and TSV files"
  WriteRegStr SHELL_CONTEXT "Software\Sheet Studio\Capabilities" "ApplicationIcon" '$\"$INSTDIR\Sheet Studio.exe$\",0'
  WriteRegStr SHELL_CONTEXT "Software\RegisteredApplications" "Sheet Studio" "Software\Sheet Studio\Capabilities"
${install.join('\n')}
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
${uninstall.join('\n')}
  DeleteRegKey SHELL_CONTEXT "Software\Classes\Applications\Sheet Studio.exe"
  DeleteRegValue SHELL_CONTEXT "Software\RegisteredApplications" "Sheet Studio"
  DeleteRegKey SHELL_CONTEXT "Software\Sheet Studio\Capabilities"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
`;
}

module.exports = { windowsFileAssociations };
