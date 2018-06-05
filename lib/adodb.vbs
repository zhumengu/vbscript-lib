function connect(DSN, username, passwd, conn)
    on error resume next

    dim rst

    if isnull(conn) or isempty(conn) then
        set conn = CreateObject("ADODB.Connection")
    end if
    ' 0 adStateClose
    ' 1 adStateOpen
    if conn.state <> 1 then
        if username ="" and passwd = "" then
            conn.Open DSN , 1, 3
        else
            conn.Open "DSN=" & DSN, username, passwd
        end if
    end if
    set connect = conn

    if err.number <> 0 then
        msgbox err.number & ": " & err.description
        err.clear
        connect = false
        exit function
    end if
end function


