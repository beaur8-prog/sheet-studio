'use strict';
const { fileAssociations } = require('../electron/file-open-requests.cjs');

// Register candidates for Open With and Default apps. Do not write extension
// defaults or Explorer UserChoice: only the user chooses the preferred app.
function windowsFileAssociations() {
  const install = [], uninstall = [];
  for (const { ext, name } of fileAssociations) {
    const progId = `SheetStudio.${ext}`;
    const legacyKey = 'Software\\Classes\\' + name;
    const ifEmpty = '${If} $0 == ""', endIf = '${EndIf}';
    // Earlier installers used the type's display name as its ProgID. Retain
    // missing aliases so hashed UserChoice entries from those versions work.
    // An existing handler belonging to another application is never replaced.
    install.push(String.raw`
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}" "" "${name}"
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}\DefaultIcon" "" '$\"$INSTDIR\Sheet Studio.exe$\",0'
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}\shell\open\command" "" '$\"$INSTDIR\Sheet Studio.exe$\" $\"%1$\"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}\Application" "ApplicationName" "Sheet Studio"
  WriteRegNone SHELL_CONTEXT "Software\Classes\.${ext}\OpenWithProgids" "${progId}"
  WriteRegNone SHELL_CONTEXT "Software\Classes\Applications\Sheet Studio.exe\SupportedTypes" ".${ext}"
  WriteRegStr SHELL_CONTEXT "Software\Sheet Studio\Capabilities\FileAssociations" ".${ext}" "${progId}"
  ReadRegStr $0 SHELL_CONTEXT "${legacyKey}\shell\open\command" ""
  ${ifEmpty}
    WriteRegStr SHELL_CONTEXT "${legacyKey}" "SheetStudioLegacyOwner" "$INSTDIR\Sheet Studio.exe"
    WriteRegStr SHELL_CONTEXT "${legacyKey}\DefaultIcon" "" '$\"$INSTDIR\Sheet Studio.exe$\",0'
    WriteRegStr SHELL_CONTEXT "${legacyKey}\shell\open\command" "" '$\"$INSTDIR\Sheet Studio.exe$\" $\"%1$\"'
  ${endIf}`);
    const ifOwned = '${If} $0 == "$INSTDIR\\Sheet Studio.exe"';
    const ifCommand = '${AndIf} $1 == \'$\\"$INSTDIR\\Sheet Studio.exe$\\" $\\"%1$\\"\'';
    uninstall.push(String.raw`
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.${ext}\OpenWithProgids" "${progId}"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\SheetStudio.${ext}"
  ReadRegStr $0 SHELL_CONTEXT "${legacyKey}" "SheetStudioLegacyOwner"
  ReadRegStr $1 SHELL_CONTEXT "${legacyKey}\shell\open\command" ""
  ${ifOwned}
  ${ifCommand}
    DeleteRegKey SHELL_CONTEXT "${legacyKey}"
  ${endIf}`);
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
