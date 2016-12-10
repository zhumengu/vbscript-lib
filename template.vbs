Option Explicit
Dim strWorkDir   
Dim strLibDir
strWorkDir = Left(WScript.ScriptFullName, instrrev(WScript.ScriptFullName,"\")-1)
strLibDir = "C:\vbs\lib\"

Include "string.vbs"
Include "util.vbs"

Sub Main()
    ' 代码从这儿开始
    debug fmt("hello %x world!!!", array("zhumeng"))
    
    If strDebug <> "" Then WScript.Echo strDebug
End Sub

'
' 函数定义
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Include(sInstFile) 
    Dim fso, f, s 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    if not fso.fileexists(sInstFile) then
        if not fso.fileexists(strLibDir & sInstFile) then
            if not fso.fileexists(strWorkDir & sInstFile) then
                exit sub
            else
                sInstFile = strWorkDir & sInstFile
            end if
        else
            sInstFile = strLibDir & sInstFile
        end if
    end if

    Set f = fso.OpenTextFile(sInstFile) 
    s = f.ReadAll 
    f.Close 
    ExecuteGlobal s 
End Sub 

Dim strDebug
Sub Debug(s)
    strDebug = strDebug & s & vbcrlf
End Sub

Call Main()


' vim: nowrap tw=0 ts=4 sw=4 sts=4
