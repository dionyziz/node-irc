#
# Node 0.35 Binary Installer
#
# Compiled using Nullsoft Installer(NSIS)
#

!include "UpgradeDLL.nsh"
!include "MUI.nsh"

!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\box-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\box-uninstall.ico"
!define MUI_COMPONENTSPAGE_SMALLDESC

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "gpl.txt"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Set languages (first is default language)
!insertmacro MUI_LANGUAGE "English"
!include english.nsh
!insertmacro MUI_LANGUAGE "German"
!include german.nsh
!insertmacro MUI_LANGUAGE "Greek"
!include greek.nsh
!insertmacro MUI_LANGUAGE "Italian"
!include italian.nsh
!insertmacro MUI_LANGUAGE "Polish"
!include polish.nsh
!insertmacro MUI_LANGUAGE "Portuguese"
!include portuguese.nsh
!insertmacro MUI_LANGUAGE "Spanish"
!include spanish.nsh
!insertmacro MUI_LANGUAGE "Swedish"
!include swedish.nsh
!insertmacro MUI_RESERVEFILE_LANGDLL
!define MUI_FINISHPAGE_RUN "$INSTDIR\node.exe"

OutFile ..\..\..\node_035.exe
Name Node
Caption "Node 0.35 Setup"
BrandingText "Nullsoft Install System"
#Icon setup.ico
WindowIcon off
BGGradient C2D6F2 5296F7
CRCCheck on
InstallDir $PROGRAMFILES\Node
XPStyle on

InstType $(type_standart)
InstType $(type_full)
InstType $(type_fullwsource)
InstType $(type_minimum)

# InstType /NOCUSTOM
InstallColors 5296F7 C2D6F2
InstProgressFlags smooth colored
AutoCloseWindow true
# ShowInstDetails show
ShowUninstDetails hide

# Sections...

Section $(sec_programmfiles) SecProgrammfiles
  SectionIn 1 2 3 4 RO
  SetOutPath $INSTDIR
  # All Files + Binary
  File ..\..\node.exe
  # File ..\..\node.exe.manifest
  SetOutPath $INSTDIR\conf
  # File ..\..\conf\node.exists
  File ..\..\conf\channels.lst
  File ..\..\conf\servers.lst
  SetOutPath $INSTDIR\doc
  File ..\..\doc\*.*
  SetOutPath $INSTDIR\downloads
  File ..\..\downloads\node.exists
  SetOutPath $INSTDIR\logs
  File ..\..\logs\node.exists
  SetOutPath $INSTDIR\misc
  File ..\..\misc\*.*
  File /r ..\..\misc\nodemenu
  File /r ..\..\misc\nodeTab
  File /r ..\..\misc\PlugInsInterface
  File /r ..\..\misc\ScriptEditor
  SetOutPath $INSTDIR\scripts
  File ..\..\scripts\*.*
  SetOutPath $INSTDIR\temp
  # File ..\..\temp\node.exists
  # SetOutPath $INSTDIR\misc
  # File ..\..\misc\nodemenu\prjNodeMenu.ocx
  SetOutPath $INSTDIR\data
  File ..\..\data\*.*
  SetOutPath $INSTDIR\data
  File /r ..\..\data\graphics
  File /r ..\..\data\html
  SetOutPath $INSTDIR\data\languages
  File ..\..\data\languages\english.lang
  File ..\..\data\languages\english.jpg
  SetOutPath $INSTDIR\data\skins
  File ..\..\data\Skins\default.skin
  File /r ..\..\data\Skins\default
  SetOutPath $INSTDIR\data
  File /r ..\..\data\dialogs
  SetOutPath $INSTDIR\data
  File /r ..\..\data\smileys
  SetOutPath $INSTDIR\data
  File /r ..\..\data\sounds
  CreateDirectory $INSTDIR\data\sounds
  SetOutPath $INSTDIR\doc
  File ..\..\doc\*.*
	# SetOutPath back to the install dir
	SetOutPath $INSTDIR
  # Default Options
  #WriteRegStr HKCU "Software\VB and VBA Program Settings\Node" Skin $INSTDIR\data\Skins\default.skin
  #WriteRegStr HKCU "Software\VB and VBA Program Settings\Node" LanguageFile $INSTDIR\data\languages\english.lang
  
  DeleteRegKey HKCU "Software\VB and VBA Program Settings\Node"

  # Install Uninstaller
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Node" "DisplayName" "Node 0.35"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Node" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteUninstaller "uninstall.exe"

  # Register Components
  !insertmacro UpgradeDLL ..\NodeMenu\prjNodeMenu.ocx $INSTDIR\Misc\NodeMenu\prjNodeMenu.ocx \
    $SYSDIR
  !insertmacro UpgradeDLL ..\nodeTab\prjNodeTab.ocx $INSTDIR\Misc\NodeTab\prjNodeTab.ocx \
    $SYSDIR
  !insertmacro UpgradeDLL ..\..\node.exe $INSTDIR\node.exe \
    $SYSDIR
	!insertmacro UpgradeDLL ..\PlugInsInterface\prjNodePlugInsInterface.dll $INSTDIR\PlugInsInterface\prjNodePlugInsInterface.dll \
    $SYSDIR
  
  # Install RunTime Files
  !insertmacro UpgradeDLL RunTime\Comcat.dll $SYSDIR\Comcat.dll \
    $SYSDIR
  !insertmacro UpgradeDLL RunTime\Msvbvm60.dll $SYSDIR\Msvbvm60.dll \
    $SYSDIR
  !insertmacro UpgradeDLL RunTime\Oleaut32.dll $SYSDIR\Oleaut32.dll \
    $SYSDIR
  !insertmacro UpgradeDLL RunTime\Olepro32.dll $SYSDIR\Olepro32.dll \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\MSWINSCK.OCX $SYSDIR\MSWINSCK.OCX \
    $SYSDIR
	WriteRegStr HKCR "Licenses\2c49f800-c2dd-11cf-9ad6-0080c7e7b78d" "" "mlrljgrlhltlngjlthrligklpkrhllglqlrk"
	!insertmacro UpgradeDLL RunTime\MSCOMCTL.OCX $SYSDIR\MSCOMCTL.OCX \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\RICHTX32.OCX $SYSDIR\RICHTX32.OCX \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\TABCTL32.OCX $SYSDIR\TABCTL32.OCX \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\msscript.ocx $SYSDIR\msscript.ocx \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\mshtml.tlb $SYSDIR\mshtml.tlb \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\TLBINF32.DLL $SYSDIR\TLBINF32.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\MSFLXGRD.OCX $SYSDIR\MSFLXGRD.OCX \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\comctl32.ocx $SYSDIR\comctl32.ocx \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\COMDLG32.OCX $SYSDIR\COMDLG32.OCX \
    $SYSDIR
	#!insertmacro UpgradeDLL RunTime\Flash.OCX $SYSDIR\Flash.OCX \
  #  $SYSDIR
	!insertmacro UpgradeDLL RunTime\AUTPRX32.DLL $SYSDIR\AUTPRX32.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\IPHLPAPI.DLL $SYSDIR\IPHLPAPI.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\MSRDO20.DLL $SYSDIR\MSRDO20.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\msvcrt.dll $SYSDIR\msvcrt.dll \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\msxml3.dll $SYSDIR\msxml3.dll \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\ODKOB32.DLL $SYSDIR\ODKOB32.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\RACREG32.DLL $SYSDIR\RACREG32.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\RDOCURS.DLL $SYSDIR\RDOCURS.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\RICHED32.DLL $SYSDIR\RICHED32.DLL \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\scrrun.dll $SYSDIR\scrrun.dll \
    $SYSDIR
	!insertmacro UpgradeDLL RunTime\VB6STKIT.DLL $SYSDIR\VB6STKIT.DLL \
    $SYSDIR

  !define UPGRADEDLL_NOREGISTER
    !insertmacro UpgradeDLL RunTime\Asycfilt.dll \
      $SYSDIR\Asycfilt.dll $SYSDIR
    !insertmacro UpgradeDLL RunTime\Stdole2.tlb $SYSDIR\Stdole2.tlb \
      $SYSDIR
  !undef UPGRADEDLL_NOREGISTER
  ;Only iease DLL count on new installation
  ;Replace myprog.exe or use another detection method
  IfFileExists $INSTDIR\node.exe skipAddSharedDLL
    Push $SYSDIR\Asycfilt.dll
    Call AddSharedDLL
		Push $SYSDIR\COMDLG32.OCX
    Call AddSharedDLL
		Push $SYSDIR\Comcat.dll
    Call AddSharedDLL
    Push $SYSDIR\comctl32.ocx
    Call AddSharedDLL
		#Push $SYSDIR\flash.ocx
    #Call AddSharedDLL
    Push $SYSDIR\MSCOMCTL.OCX
    Call AddSharedDLL
		Push $SYSDIR\MSFLXGRD.OCX
    Call AddSharedDLL
		Push $SYSDIR\mshtml.tlb
    Call AddSharedDLL
		Push $SYSDIR\msscript.ocx
    Call AddSharedDLL
    Push $SYSDIR\Msvbvm60.dll
    Call AddSharedDLL
    Push $SYSDIR\MSWINSCK.OCX
    Call AddSharedDLL
    Push $SYSDIR\Oleaut32.dll
    Call AddSharedDLL
    Push $SYSDIR\Olepro32.dll
    Call AddSharedDLL
    Push $SYSDIR\RICHTX32.OCX
    Call AddSharedDLL
    Push $SYSDIR\Stdole2.tlb
    Call AddSharedDLL
		Push $SYSDIR\TABCTL32.OCX
    Call AddSharedDLL
		Push $SYSDIR\TLBINF32.DLL
    Call AddSharedDLL
    Push $SYSDIR\prjNodeMenu.ocx
    Call AddSharedDLL
    Push $SYSDIR\prjNodeTab.ocx
    Call AddSharedDLL
		Push $SYSDIR\prjNodePlugInsInterface.dll
    Call AddSharedDLL
    Push $INSTDIR\node.exe
    Call AddSharedDLL
    Push $INSTDIR\AUTPRX32.DLL
    Call AddSharedDLL
    Push $INSTDIR\IPHLPAPI.DLL
    Call AddSharedDLL
    Push $INSTDIR\MSRDO20.DLL
    Call AddSharedDLL
    Push $INSTDIR\msvcrt.dll
    Call AddSharedDLL
    Push $INSTDIR\msxml3.dll
    Call AddSharedDLL
    Push $INSTDIR\ODKOB32.DLL
    Call AddSharedDLL
    Push $INSTDIR\RACREG32.DLL
    Call AddSharedDLL
    Push $INSTDIR\RDOCURS.DLL
    Call AddSharedDLL
    Push $INSTDIR\RICHED32.DLL
    Call AddSharedDLL
    Push $INSTDIR\scrrun.dll
    Call AddSharedDLL
    Push $INSTDIR\VB6STKIT.DLL
    Call AddSharedDLL
  skipAddSharedDLL:
SectionEnd

SubSection $(sub_sec_plugins) SubSecPlugins

  	Section $(sec_plugins_alias) SecPluginsAlias
		SectionIn 2 3
		SetOutPath $INSTDIR\data\plugins
		File ..\..\data\plugins\prjpluginalias.dll
	SectionEnd
	
	Section $(sec_plugins_textstyle) SecPluginsTextstyle
		SectionIn 2 3
		SetOutPath $INSTDIR\data\plugins
		File ..\..\data\plugins\prjplugintextstyle.dll
	SectionEnd
	
	Section $(sec_plugins_winamp) SecPluginsWinamp
		SectionIn 2 3
		SetOutPath $INSTDIR\data\plugins
		File ..\..\data\plugins\prjpluginwinamp.dll
	SectionEnd
	
	Section $(sec_plugins_source) SecPluginsSource
		SectionIn 3
		SetOutPath $INSTDIR\data\plugins\source
		File /r ..\..\data\plugins\source
	SectionEnd

SubSectionEnd

Section $(sec_sourcecode) SecSourceCode
  SectionIn 3
  SetOutPath $INSTDIR
  File ..\..\*.*
SectionEnd

Section $(sec_langfiles) SecLangfiles
  SectionIn 2 3
  SetOutPath $INSTDIR\data\languages
  File ..\..\data\languages\*.*    
SectionEnd

Section $(sec_skins) SecSkins
  SectionIn 2 3
  SetOutPath $INSTDIR\data
  File /r ..\..\Data\Skins
SectionEnd

SubSection $(sub_sec_shortcuts) SubSecShortcuts
	
Section $(sec_startmenu) SecStartmenu
  SectionIn 1 2 3 4
  CreateDirectory $SMPROGRAMS\Node
  CreateShortCut $SMPROGRAMS\Node\Node.lnk $INSTDIR\node.exe "" "" "" "" "" $(desc_Node)
	CreateShortCut "$SMPROGRAMS\Node\$(ScriptEditor).lnk" $INSTDIR\misc\scripteditor\sedit.exe "" "" "" "" "" $(desc_SEdit)
	CreateShortCut $SMPROGRAMS\Node\$(Remove).lnk $INSTDIR\uninstall.exe "" "" "" "" "" $(desc_Remove)
	CreateDirectory "$SMPROGRAMS\Node\$(Links)"
		CreateShortCut "$SMPROGRAMS\Node\$(Links)\$(HomePage).lnk" "http://node.sourceforge.net" "" "" "" "" "" $(desc_HomePage)
		CreateShortCut "$SMPROGRAMS\Node\$(Links)\$(OpenDiscussion).lnk" "http://sourceforge.net/forum/forum.php?forum_id=327488" "" "" "" "" "" $(desc_Forums)
SectionEnd

Section $(sec_desktop) SecDesktop
  SectionIn 1 2 3 4
  CreateShortCut "$DESKTOP\Node 0.35.lnk" $INSTDIR\node.exe
SectionEnd

SubSectionEnd

Section "Associate with IRC-Protocol"
	SectionIn 1 2 3 4
	WriteRegStr HKCR "irc" "" "URL:irc Protocol"
  WriteRegBin HKCR "irc" "EditFlags" "02000000"
	WriteRegStr HKCR "irc\DefaultIcon" "" "$INSTDIR\node.exe"
	WriteRegStr HKCR "irc\Shell\open\command" "" "$INSTDIR\node.exe"
	WriteRegStr HKCR "irc\Shell\open\ddeexec" "" "%1"
	WriteRegStr HKCR "irc\Shell\open\ddeexec\Application" "@" "NodeIRC"
	WriteRegStr HKCR "irc\Shell\open\ddeexec\ifexec" "" "%1"
	WriteRegStr HKCR "irc\Shell\open\ddeexec\Topic" "" "Node"
SectionEnd

Section "Uninstall"
  RMDIR /r $INSTDIR
  RMDIR /r $SMPROGRAMS\Node
  Delete "$DESKTOP\Node 0.35.lnk"

  # Remove Uninstaller from Control Panel
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Node"

  # Default Options
  #DeleteRegKey HKCU "Software\VB and VBA Program Settings\Node"

  # For VB RunTime files
  Push $SYSDIR\Asycfilt.dll
  Call un.DecrementSharedDLL
  Push $SYSDIR\Comcat.dll
  Call un.DecrementSharedDLL
  Push $SYSDIR\Msvbvm60.dll
  Call un.DecrementSharedDLL
  Push $SYSDIR\Oleaut32.dll
  Call un.DecrementSharedDLL
  Push $SYSDIR\Olepro32.dll
  Call un.DecrementSharedDLL
  Push $SYSDIR\Stdole2.tlb
  Call un.DecrementSharedDLL

SectionEnd

;Assign language strings to sections
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${SecProgrammfiles} $(DESC_SecProgrammfiles)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecSourceCode} $(DESC_SecSourceCode)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecLangfiles} $(DESC_SecLangfiles)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecSkins} $(DESC_SecSkins)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecStartmenu} $(DESC_SecStartmenu)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecDesktop} $(DESC_SecDesktop)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecPluginsAlias} $(DESC_SecPlugins_Alias)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecPluginsTextstyle} $(DESC_SecPlugins_Textstyle)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecPluginsWinamp} $(DESC_SecPlugins_Winamp)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecPluginsSource} $(DESC_SecPlugins_Source)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Function .onInit

  !insertmacro MUI_LANGDLL_DISPLAY

   # Check if user is running Windows 95
   # Node might not work on this version
   # of windows, so warn the user first
   # Based on Yazno's function, http://yazno.tripod.com/powerpimpit/
   # Updated by Joost Verburg
   # Modified by Node Devel Team

   Push $R0
   Push $R1
   
   # Check to see if we're running Win NT/Win 98
   ReadRegStr $R0 HKLM \
   "SOFTWARE\Microsoft\Windows NT\CurrentVersion" CurrentVersion

   IfErrors 0 lbl_winnt
   
   ; we are not NT
   ReadRegStr $R0 HKLM \
   "SOFTWARE\Microsoft\Windows\CurrentVersion" VersionNumber
 
   StrCpy $R1 $R0 1
   StrCmp $R1 '4.0' lbl_win32_95

   Goto lbl_done
 
   lbl_winnt:
     StrCpy $R1 $R0 1
 
     StrCmp $R1 '3' lbl_winnt_x
     StrCmp $R1 '4' lbl_winnt_x
     
     # a newer Win NT version (2000, XP, 2003...)
     # OK.

     goto lbl_done
     
     lbl_winnt_x:

     MessageBox MB_YESNO|MB_ICONQUESTION $(misc_winnt) IDYES lbl_done
       Abort
   
   lbl_win32_95:
     MessageBox MB_YESNO|MB_ICONQUESTION $(misc_win95) IDYES lbl_done
       Abort

   lbl_done:

	ClearErrors
	UserInfo::GetName
	IfErrors done
	Pop $0
	UserInfo::GetAccountType
	Pop $1
	StrCmp $1 "Admin" 0 +3
		; MessageBox MB_OK 'User "$0" is in the Administrators group'
		Goto done
	MessageBox MB_OK 'The user "$0" is not in the Administrators User Group. In order to install Node you need to have Administrator rights. Please contact your system administrator or technical support group for more information.'
		Abort

	done:

FunctionEnd

Function AddSharedDLL
  Exch $R1
  Push $R0
    ReadRegDword $R0 HKLM \
      Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1
    IntOp $R0 $R0 + 1
    WriteRegDWORD HKLM \
      Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1 $R0
   Pop $R0
   Pop $R1
 FunctionEnd

Function un.DecrementSharedDLL
  Exch $R1
  Push $R0
  ReadRegDword $R0 HKLM \
    Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1
  StrCmp $R0 "" done
    IntOp $R0 $R0 - 1
    IntCmp $R0 0 rk rk uk
    rk:
      DeleteRegValue HKLM \
        Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1
      Goto done
    uk:
      WriteRegDWORD HKLM \
        Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1 $R0
      Goto done
  done:
  Pop $R0
  Pop $R1
FunctionEnd
