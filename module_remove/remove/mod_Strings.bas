Attribute VB_Name = "mod_Strings"
Option Compare Database
Option Explicit

' ---------------------------------
' MODULE:       mod_Strings
' Level:        Framework module
' Version:      1.19
' Description:  String related functions & subroutines
' Requires:     Microsoft VBScript Regular Expressions 5.5 library for RemoveChars() oReg
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/12/2016 - 1.01 - added Unicode strings, GetUnicode()
'              *BLC, 3/16/2016 - 1.02 - added ReplaceMulti() [Uplands 2016 preseason mods]
'               BLC, 5/13/2016 - 1.02 - StringFromCodePoint()
'               BLC, 6/7/2016  - 1.03 - added InternalTrim()
'               BLC, 6/10/2016 - 1.04 - added SplitInt()
'               BLC, 6/24/2016 - 1.05 - added RemoveChars(),
'                                       replaced Exit_Function > Exit_Handler
'               BLC, 8/23/2016 - 1.06 - added ExtractString()
'               BLC, 8/30/2016 - 1.07 - added ParseString()
'               BLC, 10/25/2016 - 1.08 - added InsertSpaceBeforeCaps()
'               BLC, 2/21/2017 - 1.09 - added PadString()
' --------------------------------------------------------------------
'               BLC, 3/8/2017          added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 3/8/2017 - 1.10a - imported into invasives
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added updated version to Upland db
' --------------------------------------------------------------------
'               BLC, 3/22/2017 - 1.10 - added ReplaceMulti() from uplands & inserted missing version
' --------------------------------------------------------------------
'               BLC, 4/18/2017          re-added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 4/18/2017 - 1.11 - added updated version to Invasives db, added Requires documentation
'                                       for regular expressions
'               BLC, 6/20/2017 - 1.12 - added ReplaceTextBetween()
'               BLC, 9/13/2017 - 1.13 - documented RemoveChars requirement for Microsoft VBScript Regular Expressions 5.5 library
'               BLC, 9/14/2017 - 1.14 - organized methods, added from removed mod_Utilities:
'                                       ReplaceChars() which mirrors ReplaceChars_TSB(),
'                                       TitleCaseNameSplit(), UnderscoreNameSplit(), Capitalize()
'               BLC, 10/6/2017 - 1.15 - update documentation, add Wrap()
'               BLC, 10/18/2017 - 1.16 - add Truncate()
'               BLC, 10/19/2017 - 1.17 - added ASCII tab constant
'               BLC, 11/10/2017 - 1.18 - added Ne() similar to Nz(), SetTrace()
'               BLC, 12/27/2017 - 1.19 - added constants: uLedger, uNotebook, uPicDocument
' ---------------------------------

'---------------------
' Declarations
'---------------------
' Hex Unicode constants --> use w/ ChrW() or StringFromCodepoint() if supplementary unicode characters (codepoints) outside ChrW() range
'                           Ranges: Chr (0-255) & ChrW (-32768 - 65535), StringFromCodepoint(all)
'                           Out of Range? --> Argument exception error # 5 occurs
'                           long values are listed below
'---------------------
Public Const uSpiral = &HAA5C               '-21924 (Cham Punctuation Spiral)
Public Const uAmpersand = &H26              '38     doesn't work :(
Public Const uPlusSign = &H2B               '43     plus sign
Public Const uQuestionMark = &H3F           '63     Basic Latin Question Mark
Public Const uOn = &H7C                     '124    Vertical Line
Public Const uDoubleLessThanLeft = &HAB     '171    left-pointing double angle quotation mark <<
Public Const uDegree = &HB0                 '176    degree sign
Public Const uLineHorizontal = &H332        '818    horizontal line (Combining Low LIne)
Public Const uMu = &H3BC                    '956    microns
Public Const uDegreeSimple = &H1B80         '7040   degree (Sudanese Sign Panyecek)
Public Const uDagger = &H2020               '8224   general punctuation dagger
Public Const uRArrow = &H2192               '8594   right arrow c.f. https://en.wikipedia.org/wiki/Arrow_(symbol)
Public Const uDArrow = &H2193               '8495   down arrow
Public Const uLessThanOrEqual = &H2264      '8804
Public Const uGreaterThanOrEqual = &H2265   '8805
Public Const uHourglass = &H231B            '8987   hourglass (sand in bottom only)
Public Const uSquareFoot = &H23CD           '9165
Public Const uDoubleTriangleBlkR = &H23E9   '9193   black right-pointing double triangle
Public Const uDoubleTriangleBlkL = &H23EA   '9194   black left-pointing double triangle
Public Const uAlarmClock = &H23F0           '9200
Public Const uStopwatch = &H23F1            '9201
Public Const uTimerClock = &H23F2           '9202   timer clock
Public Const uHourglassFlowing = &H23F3     '9203   hour glass with flowing sand
Public Const uMedTriangleBlkL = &H23F4      '9204   black medium left-pointing triangle
' --- new in June 2016 (unicode 9.0 release) ----
Public Const uPowerOn = &H23FD              '9213 |
Public Const uPowerToggle = &H23FC          '9212
Public Const uPower = &H23FB                '9211
' -----------------------------------------------
Public Const uCircle1 = &H2460              '9312
Public Const uCircle2 = &H2461              '9313
Public Const uCircle3 = &H2462              '9314
Public Const uCircleR = &H24C7              '9415   circled Latin captial letter R
Public Const uTriangleBlkL = &H25C0         '9664   black left-pointing triangle
Public Const uPointerBlkL = &H25C4          '9668   black left-pointing pointer
Public Const uCircleDotted = &H25CC         '9676   dotted circle
Public Const uBullet = &H25CF               '9679
Public Const uUmbrella = &H2602             '9730
Public Const uCheckboxEmpty = &H2610        '9744
Public Const uCheckboxCheck = &H2611        '9745
Public Const uCheckboxX = &H2612            '9746
Public Const uUmbrellaRain = &H2614         '9748
Public Const uGear = &H2699                 '9881 gear w/ dot in center
Public Const uMapLighthouse = &H26EF        '9967 map symbol for lighthouse (looks like gear)
Public Const uCheck = &H2714                '10004
Public Const uSpokedAsteriskHeavy = &H273D  '10045  heavy teardrop spoked asterisk (Dingbats)
Public Const uFloretteWhite = &H2740        '10048  white florette (Dingbats)
Public Const uExclamationHeavy = &H2757     '10071 heavy exclamation point
Public Const uCircleFilled1 = &H278A        '10122
Public Const uCircleFilled2 = &H278B        '10123
Public Const uCircleFilled3 = &H278C        '10124
Public Const uPlusSignHeavy = &H2795        '10133  Heavy plus sign
Public Const uBrailleDots46 = &H2815        '10261  Braille Pattern Dots-135
Public Const uBrailleDots135 = &H2828       '10280  Braille Pattern Dots-46
Public Const uBrailleDots246 = &H282A       '10282  Braille Pattern Dots-246
Public Const uBrailleDots267 = &H2862       '10338  Braille Pattern Dots-267
Public Const uCircleBulletWhite = &H29BE    '10686  circled white bullet
Public Const uCircleBullet = &H29BF         '10687  circled bullet
Public Const uVertLineCircleAbv = &H2AEF    '10991  vertical line with circle above
Public Const uPowerOff = &H2B58             '11096  off (O), heavy circle
Public Const uLTriangle = &H2BC7            '11207  left-pointing triangle
Public Const uRTriangle = &H2BC8            '11208  right-pointing triangle
Public Const uMtn = &H30D8                  '12504  mountain (Katakana Letter He)
Public Const uMtnSun = &H30DA               '12506  mountain & sun (Katakana Letter Pe)
Public Const uRipple = &HA5BF               '42431  3x water surface W (Vai Syllable Wo)
Public Const uPerson = &H10982              '67970  simple ancient person (Meroitic Hieroglyph Letter I)
Public Const uDuck = &H10996                '67990  simple ancient duck (Meroitic Hieroglyph Letter Ka)
Public Const uWavyLines = &H10A58           '68184  Kharoshthi punctuation lines (ancient Kharosthi script)
Public Const uSheepHead = &H14485           '83077  simple sheep head (Anatolian Hieroglyph A110A)
Public Const uSpiral2 = &H169B9             '92601  Bamum Letter Phase-E Ngkaami

'--- use StringFromCodepoint() from here ---
Public Const uCircledRNegative = &H1F161    '127329 negative circled Latin capital letter r
Public Const uMtnSunrise = &H1F304          '127748 mountain sunrise
Public Const uWave = &H1F30A                '127754
Public Const uDropletBlack = &H1F322        '127778
Public Const uLightningCloud = &H1F329      '127785
Public Const uGrass = &H1F33E               '127806 grass(ear of rice)
Public Const uHerb = &H1F33F                '127807
Public Const uLeafFallen = &H1F342          '127810 fallen leaf
Public Const uLeafFluttering = &H1F343      '127811 leaf fluttering in wind
Public Const uCamping = &H1F3D5             '127957
Public Const uNatlPark = &H1F3DE            '127966 path & tree
Public Const uDesert = &H1F3DC              '127964 cactus & sun
Public Const uBlkPennant = &H1F3F1          '127985 right facing black pennant
Public Const uWhtPennant = &H1F3F2          '127986 right facing white pennant
Public Const uTag = &H1F3F7                 '127991 marking tag
Public Const uSpeech = &H1F4AC              '128172
Public Const uCow = &H1F404                 '128004
Public Const uSnail = &H1F40C               '128012
Public Const uPawPrints = &H1F43E           '128062
Public Const uEye = &H1F441                 '128065
Public Const uThumbsUp = &H1F44D            '128077 thumbs up
Public Const uThumbsDown = &H1F44E          '128078 thumbs down

Public Const uUser = &H1F464                '128100 bust in silhouette
Public Const uUsers = &H1F465               '128101 busts in silhouette
Public Const uDroplet = &H1F4A7             '128167
Public Const uPageTriCorner = &H1F4C4       '128196 page facing up
Public Const uCalendarTearOff = &H1F4C6     '128198
Public Const uChartTrendUp = &H1F4C8        '128200
Public Const uChartTrendDown = &H1F4C9      '128201
Public Const uChartBar = &H1F4CA            '128202
Public Const uClipboard = &H1F4CB           '128203
Public Const uPushpin = &H1F4CC             '128204
Public Const uPushpinRnd = &H1F4CD          '128205 round-head pushpin
Public Const uPaperclip = &H1F4CE           '128206
Public Const uRuler = &H1F4CF               '128207 straight ruler
Public Const uRulerTriangle = &H1F4D0       '128208 roofing triangle
Public Const uLedger = &H1F4D2              '128210 ledger
Public Const uNotebook = &H1F4D4            '128212 notebook w/ decorative cover
Public Const uBooks = &H1F4DA               '128218
Public Const uMemo = &H1F4DD                '128221
Public Const uInbox = &H1F4E5               '128229 inbox tray
Public Const uMagnifierLeft = &H1F50D       '128269
Public Const uMagnifierRight = &H1F50E      '128270
Public Const uLinked = &H1F517              '128279 link symbol
Public Const uCamera = &H1F4F7              '128247 camera icon
Public Const uFlashCamera = &H1F4F8         '128248 camera w/flash icon
Public Const uKey = &H1F511                 '128273 key
Public Const uLocked = &H1F512              '128274
Public Const uUnlocked = &H1F513            '128275
Public Const uWrench = &H1F527              '128295
Public Const uHammer = &H1F528              '128296
Public Const uNutAndBolt = &H1F529          '128297
Public Const uSquareButtonBlack = &H1F532   '128306
Public Const uSquareButtonWhite = &H1F533   '128307
Public Const uOneOClock = &H1F550           '128336
Public Const uTwoOClock = &H1F551           '128337
Public Const uThreeOClock = &H1F552         '128338
Public Const uFourOClock = &H1F553          '128339
Public Const uFiveOClock = &H1F554          '128340
Public Const uSixOClock = &H1F555           '128341
Public Const uSevenOClock = &H1F556         '128342
Public Const uEightOClock = &H1F557         '128343
Public Const uNineOClock = &H1F558          '128344
Public Const uTenOClock = &H1F559           '128345
Public Const uElevenOClock = &H1F55A        '128346
Public Const uTwelveOClock = &H1F55B        '128347
Public Const uOneThirty = &H1F55C           '128348
Public Const uTwoThirty = &H1F55D           '128349
Public Const uThreeThirty = &H1F55E         '128350
Public Const uFourThirty = &H1F55F          '128351
Public Const uFiveThirty = &H1F560          '128352
Public Const uSixThirty = &H1F561           '128353
Public Const uSevenThirty = &H1F562         '128354
Public Const uEightThirty = &H1F563         '128355
Public Const uNineThirty = &H1F564          '128356
Public Const uTenThirty = &H1F565           '128357
Public Const uElevenThirty = &H1F566        '128358
Public Const uTwelveThirty = &H1F567        '128359
Public Const uPencil = &H1F589              '128393
Public Const uThumbsUpRev = &H1F592         '128402 reversed thumbs up
Public Const uThumbsDownRev = &H1F593       '128403 reversed thumbs down
Public Const uFingerPointL = &H1F59C        '128412 black left pointing backhand
Public Const uPicFramed = &H1F5BB           '128443 picture w/ frame
Public Const uPicDocument = &H1F5BA         '128444 document with text & picture
Public Const uFolder = &H1F5C0              '128448
Public Const uFolderOpen = &H1F5C1          '128449
Public Const uNoteEmpty = &H1F5C5           '128453
Public Const uNotepadEmpty = &H1F5C7        '128455
Public Const uNote = &H1F5C8                '128456
Public Const uNotePage = &H1F5C9            '128457
Public Const uNotepad = &H1F5CA             '128458
Public Const uDocumentEmpty = &H1F5CB       '128459
Public Const uPageEmpty = &H1F5CC           '128460 blank page
Public Const uPagesEmpty = &H1F5CD          '128461 blank pages
Public Const uDocument = &H1F5CE            '128462
Public Const uPage = &H1F5CF                '128463 page
Public Const uPages = &H1F5D0               '128464 pages
Public Const uNotepadSpiral = &H1F5D2       '128466
Public Const uCalendarSpiral = &H1F5D3      '128467
Public Const uRefresh = &H1F5D8             '128472 clockwise left & right arrows
Public Const uCancel = &H1F5D9              '128473 X
Public Const uCancel2 = &H1F5D9             '128473
Public Const uComment = &H1F5E9             '128489 speech bubble
Public Const uDelete = &H1F5F4              '128500 script ballot X
Public Const uCheckMark = &H1F5F8           '128504 check mark
Public Const uCheckItem = &H1F5F9           '128505 checked ballot box
Public Const uPedestrian = &H1F6B6          '128694
Public Const uProhibited = &H1F6C7          '128711
Public Const uHammerWrench = &H1F6E0        '128736 crossed hammer and wrench
Public Const uIsocelesTriBlkR = &H1F782     '128898 black right pointing isoceles triangle
Public Const uRHArrow = &H1F846             '129094 heavy right arrow
Public Const uLHArrow = &H1F844             '129092 heavy left arrow
Public Const uLTriangleArrow = &H1F890      '129168 leftwards triangle arrowhead
Public Const uHandshake = &H1F91D           '129309
Public Const uLizard = &H1F98E              '129422

Public Const iTabKey = 9                    'ASCII tab key value

' ---------------------------------
'   Methods
' ---------------------------------

' ---------------------------------
'   Replacements
' ---------------------------------

' ---------------------------------
' FUNCTION:     Ne
' Description:  evaluate value and return it or an alternate value
'               similar to Nz()
' Assumptions:  -
' Parameters:   InspectValue - value to inspect (variant)
'               ZeroValue - value to return if the inspection value is either empty or 0 (variant)
' Returns:      item (string)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, November 10, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2017 - initial version
' ---------------------------------
Public Function Ne(InspectValue As Variant, ZeroValue As Variant) As Variant
On Error GoTo Err_Handler

    If Len(InspectValue) = 0 Or InspectValue = 0 Then
        Ne = ZeroValue
    Else
        Ne = InspectValue
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - Ne[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetTrace
' Description:  evaluate value to determine if it was trace and return it or an alternate value
'
' Assumptions:  "T" is the UI setting for trace values
' Parameters:   InspectValue - value to check (variant)
'               TraceValue - value to return if the inspection value is "T" (variant)
' Returns:      item (string)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, November 11, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/11/2017 - initial version
' ---------------------------------
Public Function SetTrace(InspectValue As Variant, TraceValue As Variant) As Variant
On Error GoTo Err_Handler

    SetTrace = IIf(UCase(InspectValue) = "T", TraceValue, InspectValue)
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - SetTrace[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ReplaceChars
' Description:  Replaces characters in a string
' Parameters:   strTextIn - string to work on
'               strFind - character to find
'               strReplace - character to replace with
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  Unknown, unknown
' Revisions:    Unknown - Unknown
'               BLC, 9/14/2017 - added from mod_Utilities
' ---------------------------------
Public Function ReplaceChars(strTextIn As String, strFind As String, _
    strReplace As String) As String
On Error GoTo Err_Handler

    Dim intCounter As Integer
    Dim strTmp As String
    Dim chrTmp As String * 1
    
    For intCounter = 1 To Len(strTextIn)
        chrTmp = mid$(strTextIn, intCounter)
        If chrTmp <> strFind Then
            strTmp = strTmp & chrTmp
        Else
            strTmp = strTmp & strReplace
        End If
    Next intCounter
    
    ReplaceChars = strTmp

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReplaceChars[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ReplaceString
' Description:  Replaces a substring in a string with another
' Parameters:   strTextIn - string to work on
'               strFind - string to find
'               strReplace - string to replace with
'               CaseSensitive - True for case sensitive search (default=False)
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 5/17/2006 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 9/14/2017 - ReplaceString_TSB from mod_Utilities (removed) similar
'                                with required CaseSensitive parameter
'                                removed f prefix from fCaseSensitive parameter
' ---------------------------------
Public Function ReplaceString(strTextIn As String, strFind As String, _
    strReplace As String, Optional CaseSensitive As Boolean = False) As String
On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intPos As Integer
    Dim intCaseSensitive As Integer

    ' Convert the case-sensitive boolean to the comparison constant (1=binary, 2=textual)
    intCaseSensitive = CaseSensitive + 1

    strTemp = strTextIn
    intPos = InStr(1, strTemp, strFind, intCaseSensitive)

    Do While intPos > 0
        strTemp = Left$(strTemp, intPos - 1) & strReplace & mid$(strTemp, intPos + Len(strFind))
        intPos = InStr(intPos + Len(strReplace), strTemp, strFind, intCaseSensitive)
    Loop

    ReplaceString = strTemp

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReplaceString[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ReplaceMulti
' Description:  Replaces multiple strings found in a string
' Assumptions:  params() is a string array where the search string & its replacement are within
'               the item value separated by a pipe (|) character
'                   ex:      my_original_value|my_replacement_string
'               the pipe(|) separator is used to delineate the searched & replacement strings
'                   ex:     split(params(0),"|)    yields the first term to search for & its replacement (0 index = search text, 1 index = replacement text)
' Parameters:   strOriginal - original string value
'               params() - string array for strings to replace & what to replace them with
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 16, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/16/2016  - initial version
' ---------------------------------
Public Function ReplaceMulti(strOriginal As String, Params() As String) As String
On Error GoTo Err_Handler

    Dim strNew As String, aryText() As String
    Dim i As Integer

    'default
    strNew = strOriginal

    'check all params for length, then do replacement
    For i = 0 To UBound(Params)
        
        'get search & replacement text array
        aryText = Split(Params(i), "|")
        
        'replace if original strlen is > 0 and search & replacement strings exist
        If Len(strOriginal) > 0 And UBound(aryText) = 1 Then
            strNew = Replace(strNew, aryText(0), aryText(1))
        End If
    
    Next

Exit_Handler:
    ReplaceMulti = strNew
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReplaceMulti[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ReplaceTextBetween
' Description:  Replaces text between two words (or chars) found in a string
' Assumptions:  -
' Parameters:   strOriginal - original string value
'               Word1 - word or character where replacement begins
'               Word2 - word or character where replacement ends
'               Replacement - string replacing the original text segment (optional, default = "")
' Returns:      original text with the text between Word1 and Word2 replaced by the Replacement
'               text (which defaults to an empty string)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 20, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2017  - initial version
' ---------------------------------
Public Function ReplaceTextBetween(strOriginal As String, Word1 As String, _
                                        Word2 As String, Optional Replacement As String = "") As String
On Error GoTo Err_Handler

    Dim strSegment As String, strNew As String
    Dim posWord1 As Integer, posWord2 As Integer, origLength As Integer

    'default
    strNew = strOriginal

    'find positions
    posWord1 = InStr(strOriginal, Word1)
    posWord2 = InStr(strOriginal, Word2)
    origLength = Len(strOriginal)
    
    'find text to replace
    strSegment = mid(strOriginal, posWord1, posWord2 - posWord1)
    
    'replace
    strNew = Replace(strOriginal, strSegment, Replacement)

Exit_Handler:
    ReplaceTextBetween = strNew
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReplaceTextBetween[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'   Parsing
' ---------------------------------

' ---------------------------------
' FUNCTION:     ParseString
' Description:  retrieve the item from string
' Assumptions:  -
' Parameters:   str - text to parse (string)
'               idx - location of item (integer)
'               delimiter - separator between items in string (string, optional, default = "|")
' Returns:      item (string)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, July 29, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/29/2015 - initial version
'   BLC - 8/30/2016 - moved from Tree form
'   BLC - 10/6/2017 - updated documentation
' ---------------------------------
Public Function ParseString(str As String, idx As Integer, Optional delimiter As String = "|") As String
On Error GoTo Err_Handler

    Dim Items() As String
    Dim Item As String
        
    Items() = Split(str, delimiter)
    
    If UBound(Items) + 1 > idx Then
        Item = Items(idx)
    End If
    
Exit_Handler:
    ParseString = Item
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - ParseString[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     TitleCaseNameSplit
' Description:  splits a string into separate words
' Assumptions:  -
' Parameters:   strIn - string to manipulate (string)
' Returns:      string split by underscores
' Throws:       none
' References:   none
' Source/date:  Unknown, unknown
' Adapted:      Bonnie Campbell, September 14, 2017 - for NCPN tools
' Revisions:
'   Unknown - Unknown - initial version
'   BLC - 9/14/2017 - moved from mod_Utilities (removed)
' ---------------------------------
Public Function TitleCaseNameSplit(strIn As String) As String
On Error GoTo Err_Handler

    Dim strOut As String
    Dim i As Integer
    
    For i = 1 To Len(strIn)
        If IsCapital(mid(strIn, i, 1)) Then
            Select Case i
                Case 2 To (Len(strIn) - 1)  'middle letters
                    'previous letter a capital letter? --> don't put a space before it
                    '                                      unless next letter is lowercase
                    If IsCapital(mid(strIn, i - 1, 1)) Then
                        If IsCapital(mid(strIn, i + 1, 1)) Then
                            strOut = strOut & mid(strIn, i, 1)
                        Else
                            strOut = strOut & " " & mid(strIn, i, 1)
                        End If
                    Else
                        'previous letter lowercase? --> put a space
                        strOut = strOut & " " & mid(strIn, i, 1)
                    End If
                Case 1  'first letter
                    strOut = UCase(Left(strIn, 1))
                Case Len(strIn) 'last letter
                    'previous letter was a capital? --> don't put a space
                    If IsCapital(mid(strIn, i - 1, 1)) Then
                        strOut = strOut & mid(strIn, i, 1)
                    Else
                        strOut = strOut & " " & mid(strIn, i, 1)
                    End If
            End Select
        Else
            strOut = strOut & mid(strIn, i, 1)
        End If
    Next
    
    TitleCaseNameSplit = Capitalize(Trim(strOut))
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - TitleCaseNameSplit[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     UnderscoreNameSplit
' Description:  splits a string into separate words with underscore between
' Assumptions:  -
' Parameters:   strIn - string to manipulate (string)
' Returns:      string split by capitals
' Throws:       none
' References:   none
' Source/date:  Unknown, unknown
' Adapted:      Bonnie Campbell, September 14, 2017 - for NCPN tools
' Revisions:
'   Unknown - Unknown - initial version
'   BLC - 9/14/2017 - moved from mod_Utilities (removed)
' ---------------------------------
Public Function UnderscoreNameSplit(strIn As String) As String
On Error GoTo Err_Handler

    UnderscoreNameSplit = Capitalize(ReplaceChars(strIn, "_", " "))
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - UnderscoreNameSplit[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'   Alterations
' ---------------------------------

' ---------------------------------
' FUNCTION:     ChangeDelimiter
' Description:  Replaces delimiters in an input string; default is to change double-quotes
'               to single quotes
' Parameters:   strInputText - string to work on
'               strCurrDelimiter - current delimiter in the string (default: double-quote)
'               strNewDelimiter - desired replacement delimiter (default: single-quote)
' Returns:      modified string
' Throws:       none
' References:   ReplaceString
' Source/date:  John R. Boetsch, 5/17/2006
' Revisions:    JRB, 5/17/2006
'               BLC, 4/30/2015 - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' ---------------------------------
Public Function ChangeDelimiter(strInputText As String, _
    Optional strCurrDelimiter As String = """", _
    Optional strNewDelimiter As String = "'") As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    
    ' Call the replace string function, specifying the delimiter and no case-sensitive search
    strTemp = ReplaceString(strInputText, strCurrDelimiter, strNewDelimiter)
    ChangeDelimiter = strTemp

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeDelimiter[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     Capitalize
' Description:  Capitalizes input text
' Assumptions:  -
' Parameters:   strIn - string to manipulate (string)
' Returns:      capitalized string
' Throws:       none
' References:   none
' Source/date:  Unknown, unknown
' Adapted:      Bonnie Campbell, September 14, 2017 - for NCPN tools
' Revisions:
'   Unknown - Unknown - initial version
'   BLC - 9/14/2017 - moved from mod_Utilities (removed)
' ---------------------------------
Public Function Capitalize(strIn As String) As String
On Error GoTo Err_Handler

    Dim strWorking As String
    Dim strWord As String
    Dim LastWord As Boolean
    
    strIn = Trim(strIn)
    
    Do
        If InStr(strIn, " ") = 0 Then
            LastWord = True
            strWord = Trim(strIn)
        Else
            strWord = Left(strIn, InStr(strIn, " ") - 1)
            strIn = mid(strIn, InStr(strIn, " ") + 1)
        End If
        
        Select Case strWord
            Case "id", "tsn", "nps"
                strWord = UCase(strWord)
            Case Else
                strWord = UCase(Left(strWord, 1)) & mid(strWord, 2)
        End Select
        
        strWorking = strWorking & " " & strWord
    
    Loop Until LastWord
    
    Capitalize = Trim(strWorking)
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - Capitalize[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     Wrap
' Description:  Wraps input with some sort of wrapping character
'
'               Accepts a variant & wraps with whatever character or
'               group of characters passed in the optional WrapWith argument.
'               Replaces all values equal to the WrapWith text
'               with duplicates of that character which enables wrapping
'               a text string that contains quotes.
'
'               If WrapWhat value is NULL --> function returns a empty string wrapped in quotes
'               Generally used to wrap text in quotes or date values in the #
' Assumptions:  -
' Parameters:   WrapWhat - string to manipulate (variant)
'               WrapWith - text to wrap around (string, optional, default = "")
' Returns:      wrapped string
' Throws:       none
' References:
'   Dale Fye, March 10, 2000
' Source/date:  Unknown, unknown
' Adapted:      Bonnie Campbell, September 14, 2017 - for NCPN tools
' Revisions:
'   Unknown - Unknown - initial version
'   BLC - 10/6/2017 - moved from mod_Temp
' ---------------------------------
Public Function Wrap(WrapWhat As Variant, Optional WrapWith As String = """") As String
On Error GoTo Err_Handler
        
    Wrap = WrapWith & Replace(WrapWhat & "", WrapWith, WrapWith & WrapWith) & WrapWith
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - Wrap[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'   Insertions
' ---------------------------------

' ---------------------------------
' FUNCTION:     InsertSpace
' Description:  Inserts a space between capitalized letters
' Parameters:   str - string to inspect
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  theDBguy, May 20, 2010
'               http://www.utteraccess.com/forum/Split-string-capital-le-t1945127.html
' Adapted:      Bonnie Campbell, June 17, 2014
' Revisions:    BLC, 6/17/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_String, added error handling
' ---------------------------------
Public Function InsertSpace(str As String) As String
     
    On Error GoTo Err_Handler
     
     Dim strTemp As String
     Dim strChar As String
     Dim intLen As Integer
     
     If str > "" Then
          For intLen = 1 To Len(str)
               strChar = mid(str, intLen, 1)
               If Asc(strChar) >= 65 And Asc(strChar) <= 90 Then
                    strTemp = strTemp & " " & strChar
               Else
                    strTemp = strTemp & strChar
               End If
          Next
     End If
        
     InsertSpace = strTemp
     
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - InsertSpace[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     InsertSpaceBeforeCaps
' Description:  adds a space before capitals
' Assumptions:  -
' Parameters:   strInspect - string to check (string)
' Returns:      string with spaces inserted (string)
' Throws:       none
' References:
'   Bleuspam, May 20, 2010
'   http://www.utteraccess.com/forum/Split-string-capital-le-t1945127.html
' Source/date:  Bonnie Campbell, July 29, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/25/2016 - initial version
' ---------------------------------
Function InsertSpaceBeforeCaps(strInspect As String) As String
On Error GoTo Err_Handler

    Dim strTest As String, strNew As String
    Dim i As Integer

    For i = 1 To Len(strInspect)
        strTest = mid(strInspect, i, 1)
        
        If StrComp(strTest, StrConv(strTest, vbUpperCase), vbBinaryCompare) <> 0 Then
            strNew = strNew & strTest
        Else:
            strNew = strNew & " " & strTest
        End If
    Next i
    
    InsertSpaceBeforeCaps = Trim(strNew)

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - InsertSpaceBeforeCaps[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     PadString
' Description:  Class string padding function
' Assumptions:  -
' Parameters:   src - starting string (string)
'               nSize - desired string length (integer)
'               PadChar - character to use for padding (string)
' Returns:      -
' Throws:       none
' References:
'   Reinhold Thurner, August 19, 2015
'   https://sourceforge.net/projects/exifclass/
' Source/date:  Bonnie Campbell, February 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/21/2017 - initial version
' ---------------------------------
Public Function PadString(src As String, nSize As Integer, PadChar As String) As String

    If Len(src) < nSize Then
        PadString = String(nSize - Len(src), PadChar) & src
    Else
        PadString = src
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PadString[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'   Removals
' ---------------------------------

' ---------------------------------
' FUNCTION:     InternalTrim
' Description:  Removes all spaces from string (before, inside, & after)
' Parameters:   str - string to inspect
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  B. Campbell June 7, 2016
' Adapted:      -
' Revisions:    BLC, 6/7/2016 - initial version
' ---------------------------------
Public Function InternalTrim(str As String) As String
     
    On Error GoTo Err_Handler
             
     InternalTrim = Replace(Trim(str), " ", "")
     
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - InternalTrim[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     RemoveChars
' Description:  Removes non-numeric or numeric values from a string
' Assumptions:  -
' Parameters:   strInspect - string to remove non-numerics from
'               blnNumeric - whether numbers or non-numerics are returned (boolean),
'                            (true - return numbers only, false - return non-numerics only)
' Returns:      string - numeric or non-numeric portion of string depending on blnNumeric
' Throws:       none
' Dependency:
' Requires:     Microsoft VBScript Regular Expressions 5.5 library
' References:
'   Ivan F. Moala, June 12, 2004
'   http://www.xtremevbtalk.com/archive/index.php/t-172627.html
' Source/date:  Bonnie Campbell, June 24, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 6/24/2016 - initial version
' ---------------------------------
Public Function RemoveChars(ByVal strInspect As String, blnNumeric As Boolean) As String
On Error GoTo Err_Handler:

    Dim oReg As RegExp

    Set oReg = CreateObject("vbScript.regexp")

    With oReg
        If blnNumeric Then
            .Pattern = "[^\d]+" '\d -> digit character of any length
        Else
            .Pattern = "[^\D]+" '\D -> non-digit character of any length
        End If
        .Global = True
    End With

    RemoveChars = oReg.Replace(strInspect, "")

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveNonNumerics[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function
    
' ---------------------------------
' Function:     Truncate
' Description:  String truncating function
' Assumptions:  -
' Parameters:   src - starting string (string)
'               EndLength - desired string length (integer)
'               TruncateFrom - left(L) or right(R) (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2017 - initial version
' ---------------------------------
Public Function Truncate(src As String, EndLength As Integer, TruncateFrom As String) As String
    
        Select Case TruncateFrom
            Case "R", "Right"
                Truncate = Right(src, EndLength)
            Case "L", "Left"
                Truncate = Left(src, EndLength)
            Case Else
                Truncate = src
        End Select
        
Exit_Handler:
        Exit Function
        
Err_Handler:
        Select Case Err.Number
          Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - Truncate[mod_Strings])"
        End Select
        Resume Exit_Handler
    End Function

' ---------------------------------
'   Retrieve String Info
' ---------------------------------

' ---------------------------------
' FUNCTION:     CountInString
' Description:  Counts the number of instances of character(s) in a string
' Assumptions:  -
' Parameters:   strInspect - string to check
'               strFind - string to count
' Returns:      count - number of instances strFind is found in strInspect
' Throws:       none
' References:   none
' Source/date:
'
' http://stackoverflow.com/questions/5193893/count-specific-character-occurrences-in-string
' Scott Huish, June 20, 2011
' http://www.mrexcel.com/forum/excel-questions/558568-count-occurrence-string-within-string-using-visual-basic-applications.html
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/7/2015 - initial version
'   BLC, 5/1/2015 - integrated into Invasives Reporting tool
' ---------------------------------
Public Function CountInString(ByVal strInspect As String, ByVal strFind As String) As Integer
On Error GoTo Err_Handler:
     Dim Count As Integer

    'default
    Count = 0
    
    If Len(strInspect) > 0 Then
        Count = UBound(Split(strInspect, strFind))
    End If
    
    CountInString = Count

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountInString[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     StringFromCodepoint
' Description:  Handles unicode character values beyond the ranges available to
'               Chr (0-255) & ChrW (-32768 - 65535)
'               Uses surrogate characters technique
'               Uses 2 16-bit characters to code up to &H110000 characters
'               outside the main plane of characters (the basic multilingual plane or BMP).
'               The two surrogate characters are two bunches
'               of 1024 characters coded between &HD800 and &HDBFF
'               and between &HC00 and &HDFFF.
'               Therefore only CodePoints less than &H110000 are allowed,
'               and the 2 characters to code a valid CodePoint are computed
' Assumptions:  -
' Parameters:   CodePoint - string to check
' Returns:      Desired Unicode character
' Throws:       none
' References:   none
' Source/date:
'   Ben - June 17, 2014 & user1771398 - June 18, 2014
'   http://stackoverflow.com/questions/24272148/vba-word-insertsymbol-failure-with-high-value-unicode-for-maths-symbols
'   Microsoft
'   https://msdn.microsoft.com/en-us/library/windows/desktop/dd374069(v=vs.85).aspx
' Adapted:      Bonnie Campbell, May 31, 2016 - for NCPN tools
' Revisions:
'   BLC, 5/31/2016 - initial version
' ---------------------------------
Function StringFromCodepoint(ByVal CodePoint As Long) As String
On Error GoTo Err_Handler
    If CodePoint <= &HFFFF& Then
        StringFromCodepoint = ChrW(CodePoint)
        Exit Function
    ElseIf CodePoint > &H10FFFF Or CodePoint <= 0 Then
        Err.Raise 5, "Invalid Codepoint: " & str(CodePoint)
        Exit Function
    Else
        CodePoint = CodePoint - &H10000
        Dim SurrogateLow As Long
        Dim SurrogateHigh As Long
        SurrogateLow = CodePoint And &H3FF&
        SurrogateHigh = (CodePoint - SurrogateLow) / &H400&
        StringFromCodepoint = ChrW(SurrogateHigh Or &HD800&) + ChrW(SurrogateLow Or &HDC00&)
        Exit Function
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - StringFromCodePoint[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SplitInt
' Description:  Splits array of strings which are integers into an array of integers
' Assumptions:  Array passed in is actually an array of integers
' Parameters:   strInspect - string to check
'               strDelimiter - string separator
' Returns:      string array - array of integers found in strInspect
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 10, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 6/10/2016 - initial version
' ---------------------------------
Public Function SplitInt(ByVal strInspect As String, strDelimiter As String) As Variant
On Error GoTo Err_Handler:
    Dim i As Integer
    
    If Len(strInspect) > 0 Then
        Dim ary() As String
        ary = Split(strInspect, strDelimiter)
        
        For i = LBound(ary) To UBound(ary)
            ary(i) = CInt(ary(i))
        Next
        
    End If
    
    SplitInt = ary

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SplitInt[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ExtractString
' Description:  Extracts string from within a string
' Assumptions:  -
' Parameters:   strInspect - string to extract from
'               strDelimiterA - string that is before the string to extract (1 charcter)
'               strDelimiterB - string that is after the string to extract (1 character)
' Returns:      string - portion of string between delimiters A & B
' Throws:       none
' References:
'   EIV, October 6, 2015
'   http://stackoverflow.com/questions/7293461/excel-vba-extract-text-between-2-characters
' Source/date:  Bonnie Campbell, August 23, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 8/23/2016 - initial version
' ---------------------------------
Public Function ExtractString(ByVal strInspect As String, strDelimiterA As String, strDelimiterB As String) As String
On Error GoTo Err_Handler:
    
    Dim posA As Integer, posB As Integer
    
    posA = InStrRev(strInspect, strDelimiterA)
    posB = InStrRev(strInspect, strDelimiterB)
    
    ExtractString = mid(strInspect, posA + 1, posB - posA - 1)

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExtractString[mod_Strings])"
    End Select
    Resume Exit_Handler
End Function

