
' 打开Excel工作薄文档，返回工作薄对象
' sFileName 输入参数, Excel 文档完整路径
' excel, 输出/输入参数, 当已经建立 excel 对象时传入.
' 使用示例
' dim excel, workbook
' set workbook = openWorkbook("Book2.xls", excel)
Function openWorkbook(sFileName, excel)
    dim fso

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        err.Raise 1000, "ReadFile: 文件不存在 '" & sFileName & "'."
        exit function
    end if

    set fso = nothing
    
    if isempty(excel) then
        set excel = CreateObject("Excel.Application")
    end if

    set OpenWorkbook =excel.Workbooks.Open(sFileName)
end function
    
' 保存并关闭工作薄
function closeWorkbook(wbook)
    if lcase(typename(wbook)) = "workbook" then
        wbook.save
        wbook.close
    end if

    if lcase(typename(wbook)) = "object" then
        set wbook = nothing
    end if
end function
