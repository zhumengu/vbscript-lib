Option Explicit
Include array("globalvar", "string", "util")
Include "file"

Sub Main()
    sleep 1000
    set wsh = Shell()
    set fso = FileSystemObject()
    dim a:a = randfilename("log")
    dim cmd
    strcpy cmd, "cmd /c dir > "
    strcpy cmd, env("tmp")
    strcpy cmd, "\"
    strcpy cmd, a
    exec cmd
    wsh.run fmt("notepad %x\%x", array(env("tmp"), a)), 1, true
    destroy wsh
    if not fileexists(a) then _
        abort fmt("%xquit%n脚本终止.", format("yyyy-m-d", date))
    else
        fso.delete fmt("%x\%x", array(env("tmp"), a))
    end if
    destroy fso

End Sub

'
' 函数定义
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Include(sInstFile) 

    dim strWorkDir, strLibDir
    strWorkDir = Left(WScript.ScriptFullName, instrrev(WScript.ScriptFullName,"\")-1)
    strLibDir = "C:\vbs\lib\"
    if typename(sInstFile) = "String" then
        Dim fso, f, s 
        Set fso = CreateObject("Scripting.FileSystemObject") 

        if lcase(right(sInstFile, 4)) <> ".vbs" then
            sInstFile = sInstFile & ".vbs"
        end if

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
        set fso = nothing
    elseif typename(sInstFile) = "Variant()" then
        dim a 
        for each a in sInstFile
            Include a
        next
    end if
End Sub 
Call Main()


' vim: nowrap tw=0 ts=4 sw=4 sts=4
