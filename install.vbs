Option Explicit
Dim strWorkDir   
Dim strLibDir
strWorkDir = Left(WScript.ScriptFullName, instrrev(WScript.ScriptFullName,"\")-1)
strLibDir = "C:\vbs\lib\"

Include "fmt.vbs"

Sub Main()
    ' ����������ʼ
    dim wsh
    set wsh = createobject("wscript.shell")
    wsh.regwrite "HKLM\software\classes\.vbs\shellnew\filename", "template.vbs", "REG_SZ" 
    msgbox "��ģ���ļ� template.vbs �ŵ� c:\users\administrator\templates ����"
    
    If strDebug <> "" Then WScript.Echo strDebug
End Sub

'
' ��������
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


'
