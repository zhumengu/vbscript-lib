' 遍历指定目录下所有以指定字符结尾的文件
' set d = filetree("f:\a", ".txt")
' for each elem in d
     'msgbox elem & vbtab & d.item(elem)
' next
function filetree(spath, pattern)
    set dic = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    if fso.folderexists(spath) then
        Set oFolder = fso.GetFolder(sPath)
        Set oFiles  = oFolder.Files
        For Each oFile In oFiles
            If lcase(right(oFile.Path,len(pattern))) = pattern Then
                dic.Add oFile.Path, oFile.Name 
            end if
        Next

        for each elem in oFolder.SubFolders
            dim tmp_dic
            set tmp_dic = filetree(elem.path, pattern)
            for each tmp in tmp_dic
                dic.Add tmp, tmp_dic.item(tmp)
            next
        next
    end if
    set fso      = nothing
    set filetree = dic
end function
