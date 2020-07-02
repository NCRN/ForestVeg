' =================================
' MODULE:       fw_mod_Collection
' Level:        Framework module
' Version:      1.01
' Description:  Collection functions & procedures
'
' Source/date:  Bonnie Campbell, 9/27/2017
' Revisions:    BLC, 9/27/2017 - 1.00 - initial version
'               BLC, 5/16/2019 - 1.01 - added fw_ module prefix
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------


' ---------------------------------
'  Sorting
' ---------------------------------

' ---------------------------------
' SUB:          Sort
' Description:  Sorts the given collection using the Arrays.MergeSort algorithm
'                   O(n log(n)) time
'                   O(n) space
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Sub Sort(col As Collection, Optional ByRef c As IVariantComparator)
On Error GoTo Err_Handler

    Dim a() As Variant
    Dim b() As Variant
    
    a = Collections.ToArray(col)
    
    Arrays.Sort a(), c
    
    Set col = Collections.FromArray(a())
    
Exit_Handler:
    'cleanup
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Sort[fw_mod_Collection])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ToArray
' Description:  Returns an array which exactly matches this collection.
' Note:         This function is not safe for concurrent modification.
' Assumptions:  -
' Parameters:   col - collection to change to array (collection)
'               -
' Returns:      collection transformed to variant array
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function ToArray(col As Collection) As Variant
On Error GoTo Err_Handler
    
    Dim a() As Variant
    ReDim a(0 To col.Count)
    Dim i As Long
    
    For i = 0 To col.Count - 1
        a(i) = col(i + 1)
    Next i
    
    ToArray = a()
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToArray[fw_mod_Collection])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FromArray
' Description:  Returns a Collection which exactly matches the given Array
' Note:         This function is not safe for concurrent modification.
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function FromArray(a() As Variant) As Collection
On Error GoTo Err_Handler

    Dim col As Collection
    Set col = New Collection
    Dim element As Variant
    
    For Each element In a
        col.Add element
    Next element
    
    Set FromArray = col
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FromArray[fw_mod_Collection])"
    End Select
    Resume Exit_Handler
End Function