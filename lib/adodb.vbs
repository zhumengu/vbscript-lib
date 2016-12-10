function connect(DSN, conn)
    on error resume next

    dim rst

    if isnull(conn) or isempty(conn) then
        set conn = CreateObject("ADODB.Connection")
    end if
    ' 0 adStateClose
    ' 1 adStateOpen
    if conn.state <> 1 then
        conn.Open "DSN=" & DSN , 1, 3
    end if
    set connect = conn

    if err.number <> 0 then
        msgbox err.number & ": " & err.description
        err.clear
        connect = false
        exit function
    end if
end function


