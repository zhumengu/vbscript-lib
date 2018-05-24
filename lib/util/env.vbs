
function env(str1)
    dim wsh 
    set wsh = createobject("wscript.shell")
    if left(str1, 1) <> "%" then str1 = "%" & str1 & "%"
    env = wsh.expandenvironmentstrings(str1)
end function
