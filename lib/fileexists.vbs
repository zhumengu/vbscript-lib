function fileexists(sFileName)
    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        fileexists = false
        exit function
    end if
    fileexists = true

    set fso = nothing
end function

