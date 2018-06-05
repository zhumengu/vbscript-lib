## vbs/lib 库使用说明

### adodb.vbs

    connect DSN, conn

使用 DSN 字符串建立数据库连接, 返回 ADODB.Connection 对象, conn 为一个空位, 可以作为返回值使用.

示例

    connect "MYDatabase", conn
    conn.Execute "SELECT * FROM tbl_user WHERE id=1"

### copysheet.vbs

    copySheet 源, 目标, 工作薄

在工作薄之中拷贝"源"工作表到"目标"工作表, "源" 为工作表名称, "目标" 为新工作表名称, "工作薄" 为传入的 Workbook 对象, 返回目标 worksheet 对象.

示例

    Set wks = CreateObject('Excel.Application')
    wks.open("F:\abc.xls")
    Set newSheet = copySheet("Sheet1", "SheetNew", wks)
    MsgBox newSheet.Name
    set newSheet = Nothing
    wks.Close
    set wks = Nothing

### fileexists.vbs

    fileexists filename

返回 boolean 值, 文件存在返回 true, 不存在返回 false

示例

    if fileexists("c:\autoexec.bat") then
        MsgBox "autoexec.bat 存在"
    else
        MsgBox "文件不存在"
    end if

### filetree.vbs

    filetree 目录, 过滤条件

"目录" 是起始目录,将在这个目录下遍历, "过滤条件" 文件名后几个字符, 暂时不支持 "*", "?" 通配符. 返回 Dictionary 对象

示例

    mydir = "c:\windows"
    set myfiles = filetree(mydir, ".exe")
    for each f in myfiles
        s = s & f & vbcrlf
    next
    MsgBox s


### format.vbs

    format 格式, 日期

按照"格式"返回日期

示例

    format "yyyy-mm-dd", now

### readfile.vbs

    readfile filename

返回以文件行组成的数组

示例

    dim arr

    arr = readfile("abc.txt")
    for i = lbound(arr) to ubound(arr)
      msgbox arr(i)
    next





### until.vbs

  - Dictionary() 返回 dictionary 对象.
  - FileSystemObject()  返回 fso 对象.
  - Access() 返回 Access.Application 对象.

示例

    Set acc = Access()
    Set dic = Dictionary()
    Set ex = Excel()

  - ArrayList()  返回 System.Collections.ArrayList 对象

示例

    Set al = ArrayList()
    al.Add "balabala"
    al.Sort
    al.Reverse
    al.Count
    al.Item(0)
    al.IndexOf("balabala")
    al.ToArray()

  - List

List 类

示例

    Set lst = new List
    lst.Add "abcd"
    lst.Add "efg"
    lst.Add 123
    lst.Remove 2
    MsgBox lst.Size()

    Set iterator = lst.GetIterator()
    do while iterator.hasNext()
      MsgBox iterator.GetNext()
    loop

    for i = 0 to lst.Size() - 1
      MsgBox lst.GetItem(i)
    next

    for each i in lst.GetArray()
      MsgBox i
    next

    lst.Clear()
