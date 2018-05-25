

' strSheetname �������, ����������
' sFileName �������, �������ļ�����·��
' workbook/excel ������������, û������ʵ��ʱ����������
' ʹ��ʾ��
' dim workbook, excel, sheet1
' set sheet1 = openSheetFile("������1", "c:\������.xls", workbook,excel)
Function openSheetFile(strSheetname,sFileName,workbook, excel)
    dim fso
    dim worksheet

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        err.raise 1000,"ReadFile: �ļ������� '" & sFileName & "'."
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

' workbook �������, excel.workbook ����
' excel �������, excel.application ����
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
