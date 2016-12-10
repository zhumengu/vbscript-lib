' 参数 : 文件全名
' 返回 : 文件行的数组
Function ReadFile(strFilename)
    dim fso
    dim ts
    dim tmp

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(strFilename) then
        msgbox "ReadFile: 文件不存在 '" & strFilename & "'."
        exit function
    end if

    set ts = fso.OpenTextFile(strFilename, 1)
    tmp = ts.ReadAll()

    ReadFile = split(tmp, vbNewLine)
End Function


