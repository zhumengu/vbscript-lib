' 
' 返回文件路径信息, 返回值为数组
' (驱动器, 路径, 文件名, 文件名基本,扩展名)
' dim f:f = fparse("c:\windows\system32\cmd.exe")
function fparse(fname)
    dim drive, path, filename ,basename, ext, idrv_endpos
    drive="": path="": filename="": basename="": ext=""
    idrv_endpos = instr(fname, ":")
    if idrv_endpos <> 0 then
        drive = left(fname, idrv_endpos) & "\"
    end if

    dim ipath_startpos, ipath_endpos
    ipath_startpos = instr(idrv_endpos + 2, fname, "\")
    ipath_endpos = instrrev(fname, "\") 
    if ipath_startpos <> 0 then
        if drive = "" then
            path = mid(fname, 1, ipath_endpos - len(drive) - 1) & "\"
        else
            path = mid(fname, idrv_endpos + 2, ipath_endpos - len(drive) - 1) & "\"
        end if
        filename = mid(fname, ipath_endpos + 1)
    else
        filename = mid(fname, idrv_endpos + 2)
    end if
    if path = "" and drive = "" then
        filename = fname
    end if
    if instr(filename, ".") <> 0 then
        basename = mid(filename, 1, instrrev(filename, ".") - 1)
        ext = mid(filename, instrrev(filename, ".") + 1)
    else
        basename = filename
    end if
    fparse = array(drive, path, filename, basename, ext)
end function
