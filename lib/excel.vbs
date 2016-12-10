''' include "openworkbook.vbs"
''' include "copysheet.vbs"
''' include "opensheetfile.vbs"

' ����Դ������Ϊ�¹����������¹��������
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


' ��Excel�������ĵ������ع���������
'dim excel, workbook
'set workbook = openWorkbook("Book2.xls", excel)
Function openWorkbook(sFileName, excel)
    dim fso

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        err.Raise 1, "ReadFile: �ļ������� '" & sFileName & "'."
        exit function
    end if

    set fso = nothing
    
    if isempty(excel) then
        set excel = CreateObject("Excel.Application")
    end if

    set OpenWorkbook =excel.Workbooks.Open(sFileName)
end function
    
' ���沢�رչ�����
function closeWorkbook(wbook)
    if lcase(typename(wbook)) = "workbook" then
        wbook.save
        wbook.close
    end if

    if lcase(typename(wbook)) = "object" then
        set wbook = nothing
    end if
end function


Function openSheetFile(strSheetname,sFileName,workbook, excel)
    dim fso
    dim worksheet

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        err.raise "ReadFile: �ļ������� '" & sFileName & "'."
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
            err.Raise 1, "OpenSheetFile: " & "δ�ܴ��빤���� " & typename(workbook)
            exit function
        end if
    else
        set workbook = excel.Workbooks.Open(sFileName)
    end if

    on error resume next
    if workbook.Sheets(strSheetname) then
        set OpenSheetFile = workbook.Sheets(strSheetname)
    end if

    exit function
e1:
    
    err.Raise 1, "OpenSheetFile: " & strSheetname & "������"

End Function


function closeSheetFile(workbook, excel)
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

