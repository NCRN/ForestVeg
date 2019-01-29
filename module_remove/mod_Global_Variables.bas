Attribute VB_Name = "mod_Global_Variables"
' =================================
' MODULE:       mod_Global_Variables
' Description:  Standard module for dimensioning global variables
' Source/date:  John R. Boetsch, May 2005
' Revisions:    JRB, 5/26/2006 - updated gvar names, added gvarConnected
'               JRB, 7/7/2009 - removed gvarParentForm; added gvarWritePermission,
'                   gvarHasAccessBE

Option Compare Database
Option Explicit

' Global variables
Public gvarConnected As Boolean     ' whether or not the back-end db connection is valid
Public gvarWritePermission As Boolean   ' whether or not user has write privileges to the
                                        '   back-end db
Public gvarHasAccessBE As Boolean   ' whether or not the app has one or more Access back-ends

' The following are used to refresh objects after updates in popup forms
Public gvarRefForm As Form          ' referring form object
Public gvarRefCtl As Control        ' specific control on referring form
Public gvarRefTaxonCtl As Control   ' specific taxon control
Public gvarRefContactCtl As Control ' specific contacts control
