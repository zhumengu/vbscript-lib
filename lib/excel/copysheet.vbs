' 拷贝源工作表为新工作表，返回新工作表对象
'wbook 输入参数, 工作簿对象
'source 输入参数, 字符串, 工作表名字
'dist 输入参数, 字符串, 新工作表名字
'使用示例
'dim excel, workbook
'set workbook = openWorkbook("Book2.xls", excel)
'set newSheet = copySheet("Sheet3", "abc", workbook)
function copySheet(source, dist, wbook)
    dim sheet, b
    for each sheet in wbook.Sheets
        if source = sheet.Name then
            b = true
        end if
        if dist = sheet.Name then
            set copySheet = sheet
            exit function
        end if
    next

    if b then
        set sheet = wbook.Sheets(source)
        sheet.copy null, sheet
        wbook.Sheets(source & " (2)").Name = dist
        set copySheet = wbook.Sheets(dist)
    end if

end function

