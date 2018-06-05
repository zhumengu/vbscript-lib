function columnLetter(column)
    If column > 26 then
        columnLetter = chr(int((column - 1) / 26) + 64) _
            & chr(((column - 1) mod 26) + 65)
    else
        columnLetter = chr(column + 64)
    end if
end function

function existsSheet(sheet, workbook)
    dim i
    if typename(workbook) <> "Nothing" then
        for each i in workbook.Sheets
            if i.Name = sheet then
                existsSheet = true
            end if
        next
    end if

end function
