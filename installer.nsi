!define PRODUCE_NAME "Child Wellness"
!define PRODUCT_VERSION "1.0"
!define PRODUCT_PUBLISHER "Hamilton Family Health Team"

SetCompressor lzma

Name "${PRODUCE_NAME} ${PRODUCT_VERSION}"
OutFile "ChildWellness.exe"
InstallDir "$PROGRAMFILES\Child Wellness"
ShowInstDetails show

Section -SETTINGS
	SetOutPath "$INSTDIR"
	SetOverwrite ifnewer
SectionEnd

#For removing StartMenu shortcut in Win 7
RequestExecutionLevel user

Section -Prerequisites
	SetOutPath $INSTDIR\Prerequisites
	MessageBox MB_YESNO "Install R?" /SD IDYES IDNO endR
		File "..\Prerequisites\R-3.1.0-win.exe"
		ExecWait "$INSTDIR\Prerequisites\R-3.1.0-win.exe"
		GoTo endR
	endR:
SectionEnd

Section -Main
	SetOutPath $INSTDIR
	File "BMI_Analysis.R"
	
	WriteUninstaller $INSTDIR\uninstaller.exe
	CreateShortCut "$SMPROGRAMS\Child Wellness.lnk" "$INSTDIR\BMI_Analysis.R"
	CreateShortCut "$SMPROGRAMS\Uninstall.lnk" "$INSTDIR\uninstall.exe"

	
SectionEnd
	
#Define what uninstaller will do
Section Uninstall

	Delete $INSTDIR\uninstaller.exe
	Delete $INSTDIR\BMI_Analysis.R
	Delete "$SMPROGRAMS\Child Wellness\Child Wellness.lnk"
	Delete "$SMPROGRAMS\Child Wellness\Uninstall.lnk"

SectionEnd