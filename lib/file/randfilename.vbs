' ����������ɵ��ļ���
' ext �������, �ļ���չ��
' ʹ��ʾ��
' tmp_file = randfilename("txt")
function randfilename(ext)
    randomize
    dim i, s
    for i = 0 to 15
        s = s & hex(int(rnd()*15))
    next 
    randfilename = s & "." & ext
end function
