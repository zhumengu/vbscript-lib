
' ��Excel�������ĵ������ع���������
' sFileName �������, Excel �ĵ�����·��
' excel, ���/�������, ���Ѿ����� excel ����ʱ����.
' ʹ��ʾ��
' dim excel, workbook
' set workbook = openWorkbook("Book2.xls", excel)
Function openWorkbook(sFileName, excel)
    dim fso

    set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(sFileName) then
        err.Raise 1000, "ReadFile: �ļ������� '" & sFileName & "'."
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
