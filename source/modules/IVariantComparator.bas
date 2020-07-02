Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        IVariantComparator
' Level:        Framework class
' Version:      1.00
'
' Description:  Comparison object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' References:
'  Maciej Los, April 5, 2011
'  http://www.codeproject.com/Questions/167323/Using-a-VS-Custom-Control-in-VBA-NOT-VB
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.01 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------
'======== Standard Methods ==========

' ---------------------------------
' FUNCTION:     Compare
' Description:  Compares two variants for their sort order.
'
'               IVariantComparator provides a method, compare, that imposes a
'               total ordering over a collection of variants.
'               A class that implements IVariantComparator, called a Comparator,
'               can be passed to the Arrays.sort and Collections.sort methods
'               to precisely control the sort order of the elements.
'
'               This function should exhibit several necessary behaviors:
'                 1) compare(x,y)=-(compare(y,x))   for all x,y
'                 2) compare(x,y)>= 0               for all x,y
'                 3) compare(x,y)>=0 and compare(y,z)>=0
'                    implies compare(x,z)> 0        for all x,y,z
' Assumptions:  -
' Parameters:   v1 - first object to compare (variant)
'               v2 - second object to compare (variant)
'               -
' Returns:      -1 --> v1 should be sorted ahead of v2
'               +1 --> v2 should be sorted ahead of v1
'                0 --> the two objects are of equal precedence
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function Compare(ByRef v1 As Variant, ByRef v2 As Variant) As Long
On Error GoTo Err_Handler
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Compare[IVariantComparator class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     Reverse
' Description:  Reverse comparison of two variants for their sort order.
'
' Assumptions:  -
' Parameters:   -
' Returns:      IVariantComparator object
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function Reverse(ByRef v1 As Variant, ByRef v2 As Variant) As Long
On Error GoTo Err_Handler
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Reverse[IVariantComparator class])"
    End Select
    Resume Exit_Handler
End Function