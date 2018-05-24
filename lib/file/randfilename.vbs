' 返回随机生成的文件名
' ext 传入参数, 文件扩展名
' 使用示例
' tmp_file = randfilename("txt")
function randfilename(ext)
    randomize
    dim i, s
    for i = 0 to 15
        s = s & hex(int(rnd()*15))
    next 
    randfilename = s & "." & ext
end function
