!include "${__FILEDIR__}\..\file-associations.nsh"

Var SheetStudioExistingInstall
Var SheetStudioMachineInstall

!macro preInit
  ; Windows x64 installers write their registration in the 64-bit view.
  SetRegView 64
  StrCpy $SheetStudioExistingInstall "0"
  StrCpy $SheetStudioMachineInstall ""
  ReadRegStr $R0 HKCU "${INSTALL_REGISTRY_KEY}" InstallLocation
  ${If} ${FileExists} "$R0\${APP_EXECUTABLE_FILENAME}"
    StrCpy $SheetStudioExistingInstall "1"
  ${EndIf}
  ReadRegStr $R0 HKLM "${INSTALL_REGISTRY_KEY}" InstallLocation
  ${If} ${FileExists} "$R0\${APP_EXECUTABLE_FILENAME}"
    StrCpy $SheetStudioExistingInstall "1"
    StrCpy $SheetStudioMachineInstall $R0
  ${EndIf}
!macroend

!macro customInstallMode
  ; Older helpers supplied /currentuser even for a machine installation.
  ; With no explicit target folder, repair that existing installation instead
  ; of creating a second copy under LocalAppData. Standard NSIS upgrade removal
  ; also reconciles an existing per-user copy while preserving app data.
  !insertmacro GetDParameter $R0
  ${If} $R0 == ""
  ${AndIf} $SheetStudioMachineInstall != ""
    StrCpy $isForceMachineInstall "1"
  ${EndIf}
!macroend

!macro customPageAfterChangeDir
  ; electron-builder normally appends the product name to a custom directory.
  ; Never alter the folder of an already installed app during a visible update.
  !undef MUI_PAGE_CUSTOMFUNCTION_PRE
  !define MUI_PAGE_CUSTOMFUNCTION_PRE SheetStudioInstallFilesPre
  Function SheetStudioInstallFilesPre
    ${If} $SheetStudioExistingInstall == "0"
      Call instFilesPre
    ${EndIf}
  FunctionEnd
!macroend

!macro customFinishPage
  !define MUI_FINISHPAGE_RUN
  !define MUI_FINISHPAGE_RUN_FUNCTION SheetStudioFinishRun
  !define MUI_PAGE_CUSTOMFUNCTION_SHOW SheetStudioFinishShow
  !insertmacro MUI_PAGE_FINISH

  Function SheetStudioFinishRun
    ; A fresh installation may launch from its finish page. On reinstall or
    ; update, only the waiting helper may launch after it verifies the version.
    ${If} $SheetStudioExistingInstall == "0"
      ${StdUtils.ExecShellAsUser} $0 "$launchLink" "open" ""
    ${EndIf}
  FunctionEnd

  Function SheetStudioFinishShow
    ${If} $SheetStudioExistingInstall == "1"
      SendMessage $mui.FinishPage.Run ${BM_SETCHECK} ${BST_UNCHECKED} 0
      ShowWindow $mui.FinishPage.Run ${SW_HIDE}
    ${EndIf}
  FunctionEnd

!macroend
