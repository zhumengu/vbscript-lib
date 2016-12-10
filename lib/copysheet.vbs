' 拷贝源工作表为新工作表，返回新工作表对象
'dim excel, workbook
'set workbook = openWorkbook("Book2.xls", excel)
'set newSheet = copySheet("Sheet3", "abc", workbook)
function copySheet(source, dist, wbook)

    for each sheet in wbook.Sheets
        if source = sheet.Name then
            b = true
        end if
    next

    if b then
        set sheet = wbook.Sheets(source)
        sheet.copy null, sheet
        wbook.Sheets(source & " (2)").Name = dist
        set copySheet = wbook.Sheets(dist)
    end if

end function
