
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
