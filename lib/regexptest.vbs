' ������ʽ����ƥ��
' ��һ������: ������ʽ�ַ���
' �ڶ�������: ��Ҫ���Ե��ַ���
' ����:       ƥ�䷵����
Function regexptest(patrn,str)
    Dim regEx
    Set regEx        = New RegExp
    regEx.Pattern    = patrn
    regEx.IgnoreCase = False
    regexptest       = regEx.Test(str)
    Set regEx        = nothing
End Function
'�滻ƥ���ַ���
'msgbox (ReplaceRegMatch("9","loader runner 9.0, qtp 9.0","10"))
Function ReplaceRegMatch(patrn,str,replaceStr)
    Dim regEx
    Set regEx        = New RegExp
    regEx.Pattern    = patrn
    regEx.IgnoreCase = False
    regEx.Global     = True   'false��ʱ��ֻ���滻��һ��ƥ����ַ�������Ϊtrue ����滻����ƥ����ַ���
    ReplaceRegMatch  = regEx.Replace(str,replaceStr)
End Function
'����ƥ������
'returnRegMatch "qtp .","qtp 1 qtp 2 qtp3 qtp 4"
Function ReturnRegMatch(patrn,str)
    Dim regEx,matches,match
    Set regEx        = New RegExp
    regEx.Pattern    = patrn
    regEx.IgnoreCase = true
    regEx.Global     = true  '��ȫ������
    Set matches      = regEx.Execute(str)
    For Each match in matches
        print cstr(match.firstIndex) + " " + match.value + " " + cstr(match.length)
    Next
End Function
