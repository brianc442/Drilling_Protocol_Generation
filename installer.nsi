; Primus Dental Implant Report Generator Installer
; Created with NSIS (http://nsis.sourceforge.net/)

!define PRODUCT_NAME "Primus Dental Implant Report Generator"
!define PRODUCT_VERSION "1.0.0"
!define PRODUCT_PUBLISHER "Inosys Implant"
!define PRODUCT_WEB_SITE "https://www.inosys.com"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\PrimusReportGen.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

SetCompressor lzma

; MUI Settings
!include "MUI2.nsh"

!define MUI_ABORTWARNING
!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
; License page
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
; Directory page
!insertmacro MUI_PAGE_DIRECTORY
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
!define MUI_FINISHPAGE_RUN "$INSTDIR\Primus Dental Implant Report Generator.exe"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"

; MUI end

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "Primus_Installer_v${PRODUCT_VERSION}.exe"
InstallDir "$PROGRAMFILES\Inosys\Primus Report Generator"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Section "MainSection" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  File "dist\Primus Dental Implant Report Generator.exe"
  File "icon.ico"
  File "icon.png"
  File "inosys_logo.png"
  File "Primus Implant List - Primus Implant List.csv"

  ; Create shortcuts
  CreateDirectory "$SMPROGRAMS\Inosys"
  CreateShortCut "$SMPROGRAMS\Inosys\Primus Report Generator.lnk" "$INSTDIR\Primus Dental Implant Report Generator.exe"
  CreateShortCut "$DESKTOP\Primus Report Generator.lnk" "$INSTDIR\Primus Dental Implant Report Generator.exe"
SectionEnd

Section -AdditionalIcons
  SetOutPath $INSTDIR
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateShortCut "$SMPROGRAMS\Inosys\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\Inosys\Uninstall.lnk" "$INSTDIR\uninst.exe"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\Primus Dental Implant Report Generator.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\Primus Dental Implant Report Generator.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd

; Uninstaller
Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\Primus Dental Implant Report Generator.exe"
  Delete "$INSTDIR\icon.ico"
  Delete "$INSTDIR\icon.png"
  Delete "$INSTDIR\inosys_logo.png"
  Delete "$INSTDIR\Primus Implant List - Primus Implant List.csv"

  Delete "$SMPROGRAMS\Inosys\Uninstall.lnk"
  Delete "$SMPROGRAMS\Inosys\Website.lnk"
  Delete "$SMPROGRAMS\Inosys\Primus Report Generator.lnk"
  Delete "$DESKTOP\Primus Report Generator.lnk"

  RMDir "$SMPROGRAMS\Inosys"
  RMDir "$INSTDIR"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd