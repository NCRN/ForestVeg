Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_Math
' Level:        Framework module
' Version:      1.02
' Description:  Framework-wide related mathematical values, functions & subroutines
'
' Source/date:  Bonnie Campbell, May 10, 2016
' Revisions:    BLC, 5/10/2016 - 1.00 - initial version
'               BLC, 2/20/2019 - 1.01 - added Add2Self()
'               BLC, 5/16/2019 - 1.02 - added fw_ module prefix
' =================================

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, May 10, 2016
' Adapted:      -
' Revisions:    BLC, 5/10/2016 - initial version
' ---------------------------------

'-----------------------------------------------------------------------
' Mathematical Constants
'-----------------------------------------------------------------------
Public Const TWIPS_PER_INCH As Integer = 1440
Public Const PI As Double = 3.14159265359

'-----------------------------------------------------------------------
' Methods
'-----------------------------------------------------------------------
' ---------------------------------
' FUNCTION:     Add2Self
' Description:  implements += functionality for VBA
' Assumptions:
'   Only strings and numeric VarTypes are handled
'   --------------------------------------------------
'    Constant         Value     Description
'   --------------------------------------------------
'    vbEmpty            0       Empty (uninitialized)
'    vbNull             1       Null (no valid data)
'    vbInteger          2       Integer
'    vbLong             3       Long integer
'    vbSingle           4       Single-precision floating-point number
'    vbDouble           5       Double-precision floating-point number
'    vbCurrency         6       Currency value
'    vbDate             7       Date value
'    vbString           8       String
'    vbObject           9       Object
'    vbError            10      Error value
'    vbBoolean          11      Boolean value
'    vbVariant          12      Variant (used only with arrays of variants)
'    vbDataObject       13      A data access object
'    vbDecimal          14      Decimal value
'    vbByte             17      Byte value
'    vbLongLong         20      LongLong integer (valid on 64-bit platforms only)
'    vbUserDefinedType  36      Variants that contain user-defined types
'    vbArray            8192    Array
'   --------------------------------------------------
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Behrooz Karjoo, June 14, 2017
'   https://stackoverflow.com/questions/17256153/does-the-operator-just-not-exist-in-vba
' Source/date:  Bonnie Campbell, February, 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Public Function Add2Self(ByRef x As Variant, Optional y As Variant) As Variant
On Error GoTo Err_Handler

    Select Case varType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbDecimal
            x = x + y
        Case vbString
            x = x + y
    End Select
    
    Add2Self = x
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Add2Self[fw_mod_Math])"
    End Select
    Resume Exit_Handler
End Function