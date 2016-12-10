
Function openSheetFile(strSheetname,sFileName,workbook, excel)
    dim fso
    dim worksheet

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        err.Raise 1, "ReadFile: 文件不存在 '" & sFileName & "'."
        exit function
    end if

    set fso = nothing
    
    if isEmpty(excel) then
        set excel = CreateObject("Excel.Application")
    end if

    if (not isEmpty(workbook)) then
        if  workbook is nothing then
            set workbook = excel.Workbooks.Open(sFileName)
        elseif lcase(typename(workbook)) <> "workbook" then
            err.Raise 1, "openSheetFile: " & "未能传入工作薄 " & typename(workbook)
            exit function
        end if
    else
        set workbook = excel.Workbooks.Open(sFileName)
    end if

    on error resume next
    if workbook.Sheets(strSheetname) then
        set openSheetFile = workbook.Sheets(strSheetname)
    end if

    exit function
e1:
    
    err.Raise 1, "openSheetFile: " & strSheetname & "不存在"

End Function


function CloseSheetFile(workbook, excel)
    if isEmpty(workbook) and isEmpty(excel)  then
        exit function
    end if

    if lcase(typename(workbook)) = "workbook" then
        workbook.save
        workbook.close
    end if

    if lcase(typename(workbook)) = "object" then
        set workbook = nothing
    end if

    if lcase(typename(excel)) = "application" then
        excel.quit
    end if

    if lcase(typename(excel)) = "object" then
        set excel = nothing
    end if
        
end function
