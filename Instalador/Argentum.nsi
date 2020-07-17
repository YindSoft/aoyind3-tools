;--------------------------------
;
; Instalador de Argentum Online
; creado por AlejoLp (alejolp@alejolp.com)
; y Maraxus (juansotuyo@gmail.com)
;
;--------------------------------


;--------------------------------
; Informacion basica del programa - Modificar estos strings para cada servidor

!define PRODUCT_NAME     "Argentum Online"
!define PRODUCT_VERSION  "0.12.1"

!define GAME_CLIENT_FILE "Argentum.exe"
!define GAME_MANUAL_FILE "Manual del Juego.url"
!define WEBSITE          "http://www.argentumonline.com.ar"


; Folder in which the game files are stored (relative to script)
!define GAME_FILES       "AO"


; Folder in which the dlls and ocx for the game are stored (relative to script)
!define DEPENDS_FOLDER   "dlls"


; Nombre del grupo de registros a crearse
!define AO_BASIC_REGKEY "ArgentumOnline"


!define UNINSTALLER_NAME "uninstall.exe"


; Both icons MUST have the same size and depth!
!define APP_ICON         "install.ico"
!define APP_UNINST_ICON  "uninst.ico"


;Banner displayed on top during installation. MUST be a bmp.
!define INSTALL_BANNER   "logoao.bmp"

;Banner displayed to the left at setup / unionstall. MUST be a bmp.
!define LATERAL_BANNER   "SetupAO.bmp"

!define INCLUDE_CONFIGURE_APP "1"       ;Set it to 0 if no configuration program exists
!define CONFIGURE_APP "AOSetup.exe"     ;Name of the configuration program


!define INCLUDE_AUTOUPDATER_APP "1"              ;Set it to 0 if no auto-update program exists
!define AUTOUPDATER_APP "AutoUpdateClient.exe"   ;Name of the auto-update program


!define INCLUDE_PASS_RECOVERY_APP "1"       ;Set it to 0 if no password recovery program exists
!define PASS_RECOVERY_APP "Recuperar.exe"   ;Name of the password recovery program

;--------------------------------
; De acá en más no deberías de tocar si no sabés lo que estás haciendo...


;--------------------------------
; Variables de uso frecuente

!define AO_INSTALLDIR_REGKEY "Software\${AO_BASIC_REGKEY}"
!define AO_UNISTALLER_REGKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AO_BASIC_REGKEY}"
!define AO_SM_FOLDER "${PRODUCT_NAME} ${PRODUCT_VERSION}"
!define AO_STARTMENU_FULL_DIR "$SMPROGRAMS\${AO_SM_FOLDER}"

!define GAME_LINK_FILE_NAME "${PRODUCT_NAME} ${PRODUCT_VERSION}.lnk"

!define INSTALL_DIR_REG_NAME "Install_Dir"

;--------------------------------
;Configuration

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"

OutFile "Instalador ${PRODUCT_NAME} ${PRODUCT_VERSION}.exe"

InstallDir "$PROGRAMFILES\${PRODUCT_NAME}"

;General

CRCCheck force
SetOverwrite on
AutoCloseWindow false
ShowInstDetails show
ShowUninstDetails show

SetCompressor /SOLID lzma

!include "MUI.nsh"
!include "Library.nsh"

; Para las DLLs y OCXs
Var ALREADY_INSTALLED

; Para la creación del grupo en el Menú de Inicio
Var START_MENU_FOLDER
Var MUI_TEMP

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM ${AO_INSTALLDIR_REGKEY} "${INSTALL_DIR_REG_NAME}"


;--------------------------------
;Interface Configuration

!define MUI_ICON "${APP_ICON}"
!define MUI_UNICON "${APP_UNINST_ICON}"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "${INSTALL_BANNER}"
!define MUI_WELCOMEFINISHPAGE_BITMAP "${LATERAL_BANNER}"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP "${LATERAL_BANNER}"
!define MUI_ABORTWARNING
!define MUI_UNABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${GAME_CLIENT_FILE}"
!define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\${GAME_MANUAL_FILE}"
!define MUI_FINISHPAGE_LINK "${WEBSITE}"
!define MUI_FINISHPAGE_LINK_LOCATION "${WEBSITE}"


;--------------------------------
; Pages for instalation

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "$(MUILicense)"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_STARTMENU Application $START_MENU_FOLDER
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH


;--------------------------------
; Pages for uninstalation

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH


;--------------------------------
; Languages

!insertmacro MUI_LANGUAGE "Spanish"
!insertmacro MUI_LANGUAGE "English"


;--------------------------------
;Reserve Files

;If you are using solid compression, files that are required before
;the actual installation should be stored first in the data block,
;because this will make your installer start faster.

!insertmacro MUI_RESERVEFILE_LANGDLL


;--------------------------------
; Description of each component in each language

LangString ARGENTUM_DESC ${LANG_ENGLISH} "Basic client for ${PRODUCT_NAME} ${PRODUCT_VERSION}"
LangString ARGENTUM_DESC ${LANG_SPANISH} "Cliente básico de ${PRODUCT_NAME} ${PRODUCT_VERSION}"

LangString DESKTOP_LINK_DESC ${LANG_ENGLISH} "Adds a link to ${PRODUCT_NAME} ${PRODUCT_VERSION} in the Desktop"
LangString DESKTOP_LINK_DESC ${LANG_SPANISH} "Agrega un acceso directo a ${PRODUCT_NAME} ${PRODUCT_VERSION} en el Escritorio"

;--------------------------------
; Name of components that need to be translated

LangString DESKTOP_LINK_COMPONENT ${LANG_ENGLISH} "Create a link in the Desktop"
LangString DESKTOP_LINK_COMPONENT ${LANG_SPANISH} "Crear un acceso directo en el Escritorio"

;--------------------------------
; Tanslations of links

LangString UNINSTALL_LINK ${LANG_ENGLISH} "Uninstall ${PRODUCT_NAME}.lnk"
LangString UNINSTALL_LINK ${LANG_SPANISH} "Desinstalar ${PRODUCT_NAME}.lnk"

LangString CONFIGURATION_APP_LINK ${LANG_ENGLISH} "Configure ${PRODUCT_NAME}.lnk"
LangString CONFIGURATION_APP_LINK ${LANG_SPANISH} "Configurar ${PRODUCT_NAME}.lnk"

LangString AUTO_UPDATE_LINK ${LANG_ENGLISH} "Search for updates.lnk"
LangString AUTO_UPDATE_LINK ${LANG_SPANISH} "Buscar actualizaciones.lnk"

LangString PASS_RECOVERY_APP_LINK ${LANG_ENGLISH} "Recover password.lnk"
LangString PASS_RECOVERY_APP_LINK ${LANG_SPANISH} "Recuperar contraseña.lnk"

;--------------------------------
; Licences for each language

LicenseLangString MUILicense ${LANG_ENGLISH} "license.txt"
LicenseLangString MUILicense ${LANG_SPANISH} "license-es.txt"


;--------------------------------
; Here starts the magic!

; The stuff to install
Section "${PRODUCT_NAME} ${PRODUCT_VERSION}" SEC_ARGENTUM

  SectionIn RO
  
  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ;--------------------------------------------------------------------
  ; *** Los archivos del juego ***

  File /r "${GAME_FILES}\*"

  ;--------------------------------------------------------------------
  ; Write the installation path into the registry
  WriteRegStr HKLM ${AO_INSTALLDIR_REGKEY} "${INSTALL_DIR_REG_NAME}" "$INSTDIR"
  
  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "${AO_UNISTALLER_REGKEY}" "DisplayName" "${PRODUCT_NAME}"
  WriteRegStr HKLM "${AO_UNISTALLER_REGKEY}" "UninstallString" '"$INSTDIR\${UNINSTALLER_NAME}"'
  WriteRegDWORD HKLM "${AO_UNISTALLER_REGKEY}" "NoModify" 1
  WriteRegDWORD HKLM "${AO_UNISTALLER_REGKEY}" "NoRepair" 1
  
  WriteUninstaller "${UNINSTALLER_NAME}"

  Call CreateStartMenuGroup
SectionEnd


;--------------------------------
; Optional section (can be disabled by the user)

Section "$(DESKTOP_LINK_COMPONENT)" SEC_DESKTOP_LINK

  CreateShortCut "$DESKTOP\${GAME_LINK_FILE_NAME}" "$INSTDIR\${GAME_CLIENT_FILE}" "" "$INSTDIR\${GAME_CLIENT_FILE}" 0

SectionEnd


;--------------------------------
; Uninstaller

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey HKLM "${AO_UNISTALLER_REGKEY}"
  DeleteRegKey HKLM "${AO_INSTALLDIR_REGKEY}"

  ; Remove game files
  Delete "$INSTDIR\*.*"

  ; Remove Start Menu shortcuts and folders, if any
  !insertmacro MUI_STARTMENU_GETFOLDER Application $MUI_TEMP

  Delete "$SMPROGRAMS\$MUI_TEMP\*.*"
  RMDir /r "$SMPROGRAMS\$MUI_TEMP"

  ; Remove Desktop shortcut, if any
  Delete "$DESKTOP\${GAME_LINK_FILE_NAME}"

  ; Remove directories used
  RMDir /r "$INSTDIR"

SectionEnd


;--------------------------------
; Install VB6 runtimes

Section "-Install VB6 runtimes"

  IfFileExists "$INSTDIR\${GAME_CLIENT_FILE}" 0 new_installation
    StrCpy $ALREADY_INSTALLED 1

  new_installation:

;--------------------------------
; Librerias basicas de VB6

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_NOTPROTECTED \
     "${DEPENDS_FOLDER}\msvbvm60.dll" "$SYSDIR\msvbvm60.dll" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\oleaut32.dll" "$SYSDIR\oleaut32.dll" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\olepro32.dll" "$SYSDIR\olepro32.dll" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\comcat.dll"   "$SYSDIR\comcat.dll"   "$SYSDIR"

  !insertmacro InstallLib DLL    $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\asycfilt.dll" "$SYSDIR\asycfilt.dll" "$SYSDIR"

  !insertmacro InstallLib TLB    $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\stdole2.tlb"  "$SYSDIR\stdole2.tlb"  "$SYSDIR"
 

;--------------------------------
; OCX y DLLs

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\MSINET.ocx" "$SYSDIR\MSINET.ocx" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\RICHTX32.ocx" "$SYSDIR\RICHTX32.ocx" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\CSWSK32.ocx" "$SYSDIR\CSWSK32.ocx" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\MSWINSCK.ocx" "$SYSDIR\MSWINSCK.ocx" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\dx7vb.dll" "$SYSDIR\dx7vb.dll" "$SYSDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
     "${DEPENDS_FOLDER}\MSCOMCTL.ocx" "$SYSDIR\MSCOMCTL.ocx" "$SYSDIR"

SectionEnd


;--------------------------------
; Uninstall VB6 runtimes

Section "-un.Uninstall VB6 runtimes"

;--------------------------------
; Librerias basicas de VB6

  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\msvbvm60.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\oleaut32.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\olepro32.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\comcat.dll"
  !insertmacro UnInstallLib DLL    SHARED NOREMOVE "$SYSDIR\asycfilt.dll"
  !insertmacro UnInstallLib TLB    SHARED NOREMOVE "$SYSDIR\stdole2.tlb"

;--------------------------------
; OCX y DLLs

  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\MSINET.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\RICHTX32.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\CSWSK32.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\MSWINSCK.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\dx7vb.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\MSCOMCTL.ocx"

SectionEnd


;--------------------------------
; Installer Functions

Function .onInit

  ; Make sure we request for the language
  !insertmacro MUI_LANGDLL_DISPLAY

FunctionEnd


;--------------------------------
; Uninstaller Functions

Function un.onInit

  ; Make sure we request for the language
  !insertmacro MUI_UNGETLANGUAGE

FunctionEnd


;----------------------------------------
; Create Start Menu group

Function CreateStartMenuGroup

  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application

    ;Create shortcuts
    CreateDirectory "${AO_STARTMENU_FULL_DIR}"

    CreateShortCut "${AO_STARTMENU_FULL_DIR}\$(UNINSTALL_LINK)" "$INSTDIR\${UNINSTALLER_NAME}" "" "$INSTDIR\${UNINSTALLER_NAME}" 0
    CreateShortCut "${AO_STARTMENU_FULL_DIR}\${GAME_LINK_FILE_NAME}" "$INSTDIR\${GAME_CLIENT_FILE}" "" "$INSTDIR\${GAME_CLIENT_FILE}" 0

    StrCmp ${INCLUDE_CONFIGURE_APP} "0" +2
      CreateShortCut "${AO_STARTMENU_FULL_DIR}\$(CONFIGURATION_APP_LINK)" "$INSTDIR\${CONFIGURE_APP}" "" "$INSTDIR\${CONFIGURE_APP}" 0

    StrCmp ${INCLUDE_AUTOUPDATER_APP} "0" +2
      CreateShortCut "${AO_STARTMENU_FULL_DIR}\$(AUTO_UPDATE_LINK)" "$INSTDIR\${AUTOUPDATER_APP}" "" "$INSTDIR\${AUTOUPDATER_APP}" 0

    StrCmp ${INCLUDE_PASS_RECOVERY_APP} "0" +2
      CreateShortCut "${AO_STARTMENU_FULL_DIR}\$(PASS_RECOVERY_APP_LINK)" "$INSTDIR\${PASS_RECOVERY_APP}" "" "$INSTDIR\${PASS_RECOVERY_APP}" 0

  !insertmacro MUI_STARTMENU_WRITE_END

FunctionEnd

;--------------------------------
; Section descriptions

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_ARGENTUM} "$(ARGENTUM_DESC)"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_DESKTOP_LINK} "$(DESKTOP_LINK_DESC)"
!insertmacro MUI_FUNCTION_DESCRIPTION_END
