' works like the printf-function in C.
' takes a string with format characters and an array
' to expand.
'
' the format characters are always "%x", independ of the
' type.
'
' usage example:
'   dim str
'   str = fmt( "hello, Mr. %x, today's date is %x.", Array("Miller",Date) )
'   response.Write str
function fmt( str, args )
    dim res     ' the result string.
    res = ""

    dim pos     ' the current position in the args array.
    pos = 0

    dim i
    for i = 1 to Len(str)
        ' found a fmt char.
        if Mid(str,i,1)="%" then
            if i<Len(str) then
                ' normal percent.
                if Mid(str,i+1,1)="%" then
                    res = res & "%"
                    i = i + 1

                ' expand from array.
                elseif Mid(str,i+1,1)="x" then
                    if typename(args) = "Variant()" then
                        res = res & CStr(args(pos))
                    else
                        res = res & cstr(args)
                    end if
                    pos = pos+1
                    i = i + 1

                elseif Mid(str,i+1,1)="n" then
                    res = res & vbcrlf
                    pos = pos+1
                    i = i + 1
                end if
            end if

        ' found a normal char.
        else
            res = res & Mid(str,i,1)
        end if
    next

    fmt = res
end function


sub strcpy(byref dest, src)
    dest = dest & src
end sub

