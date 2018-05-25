' 参数 : 文件全名, 文件内容

function writefile(filename, command)
    dim fso, ts
    set fso = createobject("scripting.filesystemobject")
    set ts = fso.opentextfile(filename, 2, true)
    ts.writeline command
    ts.close
end function
