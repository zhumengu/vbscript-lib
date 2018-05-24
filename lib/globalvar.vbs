dim strWorkDir, strLibDir
strWorkDir = Left(WScript.ScriptFullName, instrrev(WScript.ScriptFullName,"\")-1)
strLibDir = "C:\vbs\lib\"

dim wsh      ' wscript.shell
dim fso      ' filesystemobject
dim file     ' filesystemobject.file
dim folder   ' filesystemobject.folder
dim ts       ' textstream
dim ado      ' adodb
dim xl       ' excel application
dim dic      ' dictionary
dim args     ' wscript.args
set args = wscript.arguments
