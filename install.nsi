OutFile "install.exe"
InstallDir `c:\vbs`

Section `install vbs lib`
	SectionIn RO
	SetOutPath `$INSTDIR`
	File /r `lib`
	WriteRegStr HKLM `software\classes\.vbs\shellnew` `filename` `template.vbs`
	SetOutPath `$TEMPLATES`
	File `template.vbs`
SectionEnd

Section `Uninstall`
	RMDir /r `$INSTDIR`
	Delete `$TEMPLATES\template.vbs`
	DeleteRegKey HKLM `software\classes\.vbs\shellnew`
SectionEnd
