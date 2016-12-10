' 正则表达式测试匹配
' 第一个参数: 正则表达式字符串
' 第二个参数: 需要测试的字符串
' 返回:       匹配返回真
Function regexptest(patrn,str)
    Dim regEx
    Set regEx        = New RegExp
    regEx.Pattern    = patrn
    regEx.IgnoreCase = False
    regexptest       = regEx.Test(str)
    Set regEx        = nothing
End Function
'替换匹配字符串
'msgbox (ReplaceRegMatch("9","loader runner 9.0, qtp 9.0","10"))
Function ReplaceRegMatch(patrn,str,replaceStr)
    Dim regEx
    Set regEx        = New RegExp
    regEx.Pattern    = patrn
    regEx.IgnoreCase = False
    regEx.Global     = True   'false的时候只会替换第一个匹配的字符串。若为true 则会替换所有匹配的字符串
    ReplaceRegMatch  = regEx.Replace(str,replaceStr)
End Function
'返回匹配内容
'returnRegMatch "qtp .","qtp 1 qtp 2 qtp3 qtp 4"
Function ReturnRegMatch(patrn,str)
    Dim regEx,matches,match
    Set regEx        = New RegExp
    regEx.Pattern    = patrn
    regEx.IgnoreCase = true
    regEx.Global     = true  '打开全局搜索
    Set matches      = regEx.Execute(str)
    For Each match in matches
        print cstr(match.firstIndex) + " " + match.value + " " + cstr(match.length)
    Next
End Function
