' ���� : �ļ�ȫ��, �ļ�����

function writefile(filename, command)
    dim fso, ts
    set fso = createobject("scripting.filesystemobject")
    set ts = fso.opentextfile(filename, 2, true)
    ts.writeline command
    ts.close
end function
