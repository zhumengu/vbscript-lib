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

sub destroy(obj)
    set obj = nothing
end sub

function ArrayList()
    Set Arraylist = CreateObject("System.Collections.ArrayList")
'list.Add "Banana"
'list.Add "Apple"
'list.Add "Pear"

'list.Sort
'list.Reverse

'wscript.echo list.Count                 ' --> 3
'wscript.echo list.Item(0)               ' --> Pear
'wscript.echo list.IndexOf("Apple", 0)   ' --> 2
'wscript.echo join(list.ToArray(), ", ") ' --> Pear, Banana, Apple
end function

Class List
  Private mArray

  Private Sub Class_Initialize()
    mArray = Empty
  End Sub

  ' Appends the specified element to the end of this list.
  Public Sub Add(element)
    If IsEmpty(mArray) Then
      ReDim mArray(0)
      mArray(0) = element
    Else
      If mArray(UBound(mArray)) <> Empty Then
        ReDim Preserve mArray(UBound(mArray)+1)        
      End If
      mArray(UBound(mArray)) = element
    End If
  End Sub

  '  Removes the element at the specified position in this list.
  Public Sub Remove(index)
    ReDim newArray(0)
    For Each atom In mArray
      If atom <> mArray(index) Then
        If newArray(UBound(newArray)) <> Empty Then
          ReDim Preserve newArray(UBound(newArray)+1)
        End If
        newArray(UBound(newArray)) = atom
      End If
    Next
    mArray = newArray
  End Sub

  ' Returns the number of elements in this list.
  Public Function Size
    Size = UBound(mArray)+1
  End Function

  ' Returns the element at the specified position in this list.
  Public Function GetItem(index)
    GetItem = mArray(index)
  End Function

  ' Removes all of the elements from this list.
  Public Sub Clear
    mArray = Empty
  End Sub

  ' Returns true if this list contains elements.
  Public Function HasElements
    HasElements = Not IsEmpty(mArray)
  End Function

  Public Function GetIterator
    Set iterator = New ArrayIterator
    iterator.SetArray = mArray
    set GetIterator = iterator
  End Function

  Public Function GetArray
    GetArray = mArray
  End Function

End Class

Class ArrayIterator
  Private mArray
  Private mCursor  

  Private Sub Class_Initialize()
    mCursor = 0
  End Sub

  Public Property Let SetArray(array)
    mArray = array    
  End Property

  Public Function HasNext
    HasNext = (mCursor < UBound(mArray)+1)
  End Function

  Public Function GetNext
    GetNext = mArray(mCursor)
    mCursor = mCursor + 1
  End Function
End Class
'
