function format(interval, aDate)
    set regEx = new RegExp
    regEx.IgnoreCase = true
    regEx.Global = true

    format = interval

    const PATTERNS = "yyyy,q,m,y,d,ww,w,h,n,s"
    for each ptrn in split(PATTERNS, ",") 
        regEx.Pattern = ptrn
        format = regEx.Replace(format, datepart(ptrn, aDate))
    next
end function
