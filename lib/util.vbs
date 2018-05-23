Include "util\list.vbs"
Include "util\env.vbs"
Include "util\format.vbs"
Include "util\hide.vbs"

sub abort(msg)
    msgbox msg
    wscript.quit
end sub

sub sleep(m)
    wscript.sleep m
end sub

sub exec(exe)
    Shell().run exe, 0, true
end sub


sub destory(obj)
    set obj = nothing
end sub

function Dictionary()
    set Dictionary = createobject("scripting.dictionary")
end function

function FileSystemObject()
    set Filesystemobject = createobject("scripting.filesystemobject")
end function

function Access()
    set Access = createobject("access.application")
end function

function Excel()
    set Excel = createobject("excel.application")
end function

function Shell()
    set Shell = createobject("Wscript.Shell")
end function

'list.Add "Banana"
'list.Add "Apple"
'list.Add "Pear"

'list.Sort
'list.Reverse

'wscript.echo list.Count                 ' --> 3
'wscript.echo list.Item(0)               ' --> Pear
'wscript.echo list.IndexOf("Apple", 0)   ' --> 2
'wscript.echo join(list.ToArray(), ", ") ' --> Pear, Banana, Apple
function ArrayList()
    Set Arraylist = CreateObject("System.Collections.ArrayList")
end function
