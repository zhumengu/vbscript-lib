' ���� : �ļ�ȫ��
' ���� : �ļ��е�����
Function ReadFile(strFilename)
    dim fso
    dim ts
    dim tmp

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(strFilename) then
        msgbox "ReadFile: �ļ������� '" & strFilename & "'."
        exit function
    end if

    set ts = fso.OpenTextFile(strFilename, 1)
    tmp = ts.ReadAll()

    ReadFile = split(tmp, vbNewLine)
End Function


