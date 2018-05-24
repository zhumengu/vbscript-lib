
sub hide()
    dim wsh
    set wsh = createobject("wscript.shell")
    if lcase(right(host, len(host) - instrrev(host, "\"))) = "wscript.exe" then
        wsh.run "cscript """ & wscript.scriptfullname & chr(34), 0
        wscript.quit
    end if
end sub
