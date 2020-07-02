Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_BatchImport
' Level:        Framework module
' Version:      1.01
' Description:  Import functions & procedures
'
' Source/date:  Bonnie Campbell, 6/29/2016
' Revisions:    BLC, 6/29/2016 - 1.00 - initial version
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
' SUB:          BatchImportImagesToDb
' Description:  Import all photos to the database.
' Assumptions:  Folder & files will remain as long as the user doesn't delete them
' Parameters:   DirPath - directory full path (string
' Returns:      -
' Throws:       none
' References:   none
' Requires:     -
' Source/date:
'   HK1, March 9, 2011
'   http://stackoverflow.com/questions/5238299/importing-images-into-ms-access-using-vba
' Adapted:      Bonnie Campbell, June 29, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/29/2016  - initial version
' ---------------------------------
Public Sub BatchImportImagesToDb(DirPath As String)
On Error GoTo Err_Handler
    
    Dim FileName As String
    Dim x As String
    
    FileName = dir(DirPath)

    Do Until FileName = ""
        Select Case LCase(Right(FileName, 4))
            Case ".jpg" ', ".gif", ".bmp"
            
                'Photo record:  (* = reqd)
                ' *PhotoDate, *PhotoType, *Photographer_ID, DigitalFilename,
                ' NCPNImageID, PhotogFacing, PhotogLocation, PhotogLocationDescr,
                ' PhotogOrientation, SurveyPoint_ID,
                ' SubjectLocation,
                ' IsCloseup, IsReplacement, IsSkipped, InActive
                ' *LastPhotoUpdate,
                ' *CreateDate, *CreatedBy_ID, *LastModified, *LastModifiedBy_ID
                
                
                'https://msdn.microsoft.com/en-us/library/windows/desktop/ms630826(v=vs.85).aspx#SharedSample012
                'added reference: Microsoft Windows Image Acquisition Library v2.0
                ' C:\WINDOWS\System32\wiaaut.dll
                
Dim img 'As ImageFile
Dim s 'As String
Dim v 'As Vector

Set img = CreateObject("WIA.ImageFile")

img.LoadFile "C:\WINDOWS\Web\Wallpaper\Autumn.jpg"

s = "Width = " & img.Width & vbCrLf & _
    "Height = " & img.Height & vbCrLf & _
    "Depth = " & img.PixelDepth & vbCrLf & _
    "HorizontalResolution = " & img.HorizontalResolution & vbCrLf & _
    "VerticalResolution = " & img.VerticalResolution & vbCrLf & _
    "FrameCount = " & img.FrameCount & vbCrLf

If img.IsIndexedPixelFormat Then
    s = s & "Pixel data contains palette indexes" & vbCrLf
End If

If img.IsAlphaPixelFormat Then
    s = s & "Pixel data has alpha information" & vbCrLf
End If

If img.IsExtendedPixelFormat Then
    s = s & "Pixel data has extended color information (16 bit/channel)" & vbCrLf
End If

If img.IsAnimated Then
    s = s & "Image is animated" & vbCrLf
End If

If img.Properties.Exists("40091") Then
    Set v = img.Properties("40091").Value
    s = s & "Title = " & v.String & vbCrLf
End If

If img.Properties.Exists("40092") Then
    Set v = img.Properties("40092").Value
    s = s & "Comment = " & v.String & vbCrLf
End If

If img.Properties.Exists("40093") Then
    Set v = img.Properties("40093").Value
    s = s & "Author = " & v.String & vbCrLf
End If

If img.Properties.Exists("40094") Then
    Set v = img.Properties("40094").Value
    s = s & "Keywords = " & v.String & vbCrLf
End If

If img.Properties.Exists("40095") Then
    Set v = img.Properties("40095").Value
    s = s & "Subject = " & v.String & vbCrLf
End If

MsgBox s

                
            
            Case Else
                'Ignore other file extentions
        End Select
        FileName = dir 'Get next file
    Loop
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - BatchImportImages[fw_mod_BLOB])"
    End Select
    Resume Exit_Handler
End Sub