Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_App_Settings
' Level:        Application module
' Version:      1.16
' Description:  Application-wide related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015  - 1.00 - initial version
'               BLC, 11/20/2015 - 1.01 - added priority & status icons
'               BLC, 6/7/2016   - 1.02 - updated documentation & added ACCESS_ROLES (Big Rivers App)
'               BLC, 6/20/2016  - 1.03 - added DB_ADMIN_FORM and documentation
'               BLC, 9/1/2016   - 1.04 - updated APP_SYS_TABLES
'               BLC, 9/7/2016   - 1.05 - added LINK_NORMAL_TEXT for disabling links
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added updated version to Upland db
' --------------------------------------------------------------------
'               BLC, 3/22/2017  - 1.06 - removed big rivers only components
'                                        revised for uplands
' --------------------------------------------------------------------
'               BLC, 4/24/2017          added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 5/3/2017  - 1.08 - added VCS_FULL_PATH for running VCS functions/subroutines
'               BLC, 7/5/2017  - 1.09 - added QUADRATS_PER_TRANSECT to make
'                                       adding quadrats for new transects flexible in case
'                                       # changes from 3 quadrats per transect
'               BLC, 7/12/2017 - 1.10 - added VCS_SAVE_TABLES for tables to backup (lookups)
'               BLC, 7/28/2017 - 1.11 - changed DEV_MODE to global variable vs. constant to
'                                       allow user to set via tglDevMode toggle control in UI
' --------------------------------------------------------------------
'               BLC, 9/7/2017  - 1.12 - merged common code for framework from Upland, Invasives, Big Rivers dbs
' --------------------------------------------------------------------
'                     BLC, 6/15/2017 - 1.07 - merged prior version w/ current
' --------------------------------------------------------------------
'                               BLC, 5/1/2015 - 1.01 - added DEV_MODE constant
'                               BLC, 5/13/2015 - 1.02 - added UI enabled/disabled color constants
'                               BLC, 5/19/2015 - 1.03 - added FIX_LINKED_DBS flag constant
'                               BLC, 5/28/2015 - 1.04 - added MAIN_APP_MENU constant
' --------------------------------------------------------------------
'                       BLC, 6/19/2017 - 1.08 - added APP_RELEASE_ID constant value for
'                                               2017 Pre-Season Invasives Reporting Tool (tsys_App_Releases)
'                       BLC, 6/26/2017 - 1.09 - added REMOVE_RESULT_TABLES constant
' --------------------------------------------------------------------
'               BLC, 11/3/2017 - 1.13 - added g_ModalSedimentSizeIDs
'               BLC, 11/6/2017 - 1.14 - added APP to identify protocol application
'               BLC, 10/3/2018 - 1.15 - added APP_SUBSET to distinguish sub-protocols,
'                                       DB_DISTRIBUTED to identify databases sent to collaborators
'               BLC, 5/16/2019 - 1.16 - added fw_ module prefix
' =================================

' ---------------------------------
' GLOBALS:      global values set for application
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, June 2016
' Adapted:      -
' Revisions:    BLC, 6/6/2016 - initial version (NCPN WQ Big Rivers App, App_Templates)
'               BLC, 11/3/2017 - added g_ModalSedimentSizeIDs
'               BLC, 11/6/2017 - added APP to identify protocol application
'               BLC, 10/3/2018 - added APP_SUBSET to distinguish sub-protocols,
'                                DB_DISTRIBUTED to identify databases sent to collaborators
' ---------------------------------
'Public g_AppTemplates As Scripting.Dictionary     'global dictionary for application templates (if any)

Public APP As String                               'global setting for protocol application
Public APP_SUBSET As String                        'global setting for protocol subset (grassland, forest, camp)

Public DB_DISTRIBUTED As Boolean                   'global setting for if current db file is passed to collaborators
                                                   '(if so, DEV_MODE, ADMIN buttons should be hidden)
                                                   
Public gSubReportCount As Integer                  'global counter for subreports
Public g_ModalSedimentSizeIDs As Scripting.Dictionary 'global dictionary for modal sediment size ID #s
                                                     '(from AppEnum)

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, May 2014
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version (NCPN WQ Utilities Tool, WATER_YEAR_START & WATER_YEAR_END)
'               BLC, 4/22/2015 - adapted to generic tools (NCPN Invasives Reporting Tool) by adding
'                                USER_ACCESS_CONTROL (False - gives users full control in apps w/o controls,
'                                                     True - relies on user access control settings)
'                                DB_SYS_TABLES & APP_SYS_TABLES (handle table arrays for the database/
'                                   application)
'               BLC, 4/30/2015 - add DB_ADMIN_CONTROL flag to handle applications w/o full DbAdmin subform & controls
'                                add MAIN_APP_FORM constant to handle applications where frm_Switchboard is NOT the main form
'                                add APP_RELEASE_ID constant to handle application release ID w/o full DbAdmin subfrom & controls
'               BLC, 5/1/2015  - add DEV_MODE constant to enable menus typically off during use
'               BLC, 5/13/2015 - shifted UI enable/disabled colors from TempVars set in initialize (mod_App_UI) to constants
'               BLC, 5/19/2015 - added FIX_LINKED_DBS flag to handle applications which require updates of tbl_Dbs via FixLinkedDb
'                                (usually when DbAdmin is not fully implemented)
'               BLC, 5/28/2015 - added MAIN_APP_MENU to handle applications w/ main menu forms (not tabbed switchboards)
'               BLC, 4/4/2016  - added LOCATION_TYPES to allow specific types only, RECORD_ACTIONS, CONTACT_ROLES, PARKS
'               BLC, 6/7/2016  - added ACCESS_ROLES to set user application permissions
'               BLC, 9/7/2016  - added LINK_NORMAL_TEXT & _BKGD for disabling tile links
'               BLC, 5/3/2017  - added VCS_FULL_PATH for running VCS functions/subroutines
'               BLC, 7/28/2017 - changed DEV_MODE to global variable vs. constant to
'                                allow user to set via tglDevMode toggle control in UI
' --------------------------------------------------------------------
'               BLC, 9/7/2017  - merged common code for framework from Upland, Invasives, Big Rivers dbs
' --------------------------------------------------------------------
'               BLC, 6/15/2017 - merged w/ prior version
'               BLC, 6/19/2017 - added APP_RELEASE_ID constant value for 2017 Pre-Season Invasives Reporting Tool (tsys_App_Releases)
'               BLC, 6/26/2017 - added REMOVE_RESULT_TABLES constant for 2017 Pre-Season Invasives
' --------------------------------------------------------------------
'               BLC, 10/18/2017 - added CREATE_ENUMS for turning ON/OFF enum creation from enum table
' ---------------------------------
                                                                'Version Control System (VCS) db (contains modules for version control)
                                                                'Tables to save for VCS (e.g. lookups)
Public Const VCS_SAVE_TABLES As String = "tlu_projects, " & _
    "tlu_Cover_Class, tlu_Crown_Class, tlu_Disturbance, " & _
    "tlu_Eco_Site, tlu_Effervescence, tlu_Hillslope_Position, " & _
    "tlu_LP_Disturbance, tlu_LP_Soil_Surface, tlu_Monument_Code, " & _
    "tlu_NCPN_Plants, tlu_Parks, tlu_Profile_Depth, " & _
    "tlu_Rock_Frag_Q, tlu_Rock_Frag_Size, tlu_Sand_Modifier, " & _
    "tlu_Slope_Shape, tlu_Soil_Depth, tlu_Soil_Survey_Area, " & _
    "tlu_Soil_Texture, tlu_Veg_Type"
Public Const USER_ACCESS_CONTROL As Boolean = True             'Boolean flag -> db includes user access control or not
Public Const DB_ADMIN_CONTROL As Boolean = False                'Boolean flag -> db does not include DbAdmin subform & controls
Public Const FIX_LINKED_DBS As Boolean = False                  'Boolean flag -> db requires tbl_Dbs to be updated via FixLinkedDb (usually when DbAdmin is not fully implemented)
Public Const MAIN_APP_FORM As String = "Main"                   'String -> main tabbed form (frm_Switchboard, etc.)
Public Const MAIN_APP_MENU As String = "Main"                   'String -> main tabbed form (frm_Switchboard, etc.)
Public Const APP_RELEASE_ID As String = ""                      'String -> release ID (tsys_App_Release.Release_ID) for current release
                                                                '          used when db doesn't include full DbAdmin subform & controls, otherwise NULL
Public Const APP_URL As String = "science.nature.nps.gov/im/units/ncpn/datamanagement.cfm"
                                                                'String -> website URL for application
                                                                '          used when db doesn't include full DbAdmin subform & controls, otherwise NULL
Public DEV_MODE As Boolean                                      'Boolean flag -> enable menus, show controls when typically they'd be OFF
                                                                '        flag is set via DEV_MODE toggle in UI

Public CREATE_ENUMS As Boolean                                  'Boolean flag -> turns ON/OFF enum creation
                                                                '                from enum table
Public Const ACCESS_ROLES As String = "admin,power user,data entry,read only"
                                                                'String -> used in setting user application access level & permissions
'Public Const SWITCHBOARD As String = "Main"                     'String -> main application form
Public Const DB_ADMIN_FORM As String = "DbAdmin"                'String -> main db administrative form
Public Const BACKEND_REQUIRED As Boolean = True                 'Boolean flag -> identifies if back-end required
Public Const REMOVE_RESULT_TABLES As Boolean = True             'Boolean flag -> clears species by route tables for TCount, PctCover, SE, & related
                                                                '                related tables, Park_VisitYr_SpeciesCover_by_Route_Result
                                                                '                is left alone; if False all tables are left alone
                                                                '                & are regenerated &  overwritten when the report is run again for that park/year

'-----------------------------------------------------------------------
' Database Type
'-----------------------------------------------------------------------
Public Const BACKEND_TYPE As String = "ACCESS"

'-----------------------------------------------------------------------
' Database System Tables
'-----------------------------------------------------------------------
'   Array("App_Defaults", "BE_Updates", "Link_Dbs", "Link_Tables")
'   tsys_App_Defaults -> default application settings
'   tsys_BE_Updates   -> updates to post to remot back-end copies
'   tsys_Link_Dbs     -> info about linked back-end dbs
'   tsys_Link_Tables  -> info about linked tables
'-----------------------------------------------------------------------
' Application Backend System Tables
'-----------------------------------------------------------------------
'   Array("App_Releases", "Bug_Reports", "Logins", "User_Roles")
'   tsys_App_Releases -> list of application releases
'   tsys_Bug_Reports  -> tracking for known issues
'   tsys_Logins       -> system use monitoring
'   tsys_User_Roles   -> assign user access priviledges  [deprecated to Contact_Access]
'-----------------------------------------------------------------------
' SEE ALSO >>>> SysTablesExist() function
'-----------------------------------------------------------------------
Public Const DB_SYS_TABLES As String = "App_Defaults, Link_Files, Link_Tables"
Public Const APP_SYS_TABLES As String = "App_Releases, Bug_Reports, Logins"

'-----------------------------------------------------------------------
' User Interface Colors
'-----------------------------------------------------------------------
'std control colors
Public Const CTRL_DISABLED As Long = lngLtGray
Public Const CTRL_ADD_ENABLED As Long = lngLime
Public Const CTRL_REMOVE_ENABLED As Long = lngLtOrange
Public Const TEXT_ENABLED As Long = lngBlue
Public Const TEXT_DISABLED As Long = lngGray

'highlight text for tile links
Public Const LINK_HIGHLIGHT_TEXT As Long = lngBlue
Public Const LINK_HIGHLIGHT_BKGD As Long = lngYelLime
Public Const HIGHLIGHT_MISSING_VALUE As Long = lngYellow
Public Const LINK_NORMAL_TEXT As Long = lngGray50

Public Const PROGRESS_BAR As Long = lngLime

'-----------------------------------------------------------------------
' Icons
'-----------------------------------------------------------------------
Public Const ICON_PATH As String = "Z:\_____LIB\dev\git_projects\icons\small\"

Public Const FLAG_RED As String = ICON_PATH & "flag_red" & ".png"
Public Const FLAG_LIME As String = ICON_PATH & "flag_lime" & ".png"
Public Const FLAG_ORANGE As String = ICON_PATH & "flag_orange" & ".png"
Public Const FLAG_LTBLUE As String = ICON_PATH & "flag_ltblue" & ".png"
Public Const FLAG_BLUE As String = ICON_PATH & "flag_blue" & ".png"
Public Const FLAG_NAVY As String = ICON_PATH & "flag_navy" & ".png"
Public Const FLAG_PURPLE As String = ICON_PATH & "flag_purple" & ".png"

Public Const DOT_RED As String = ICON_PATH & "dot_red" & ".png"
Public Const DOT_LIME As String = ICON_PATH & "dot_lime" & ".png"
Public Const DOT_ORANGE As String = ICON_PATH & "dot_orange" & ".png"
Public Const DOT_LTBLUE As String = ICON_PATH & "dot_ltblue" & ".png"
Public Const DOT_BLUE As String = ICON_PATH & "dot_blue" & ".png"
Public Const DOT_NAVY As String = ICON_PATH & "dot_navy" & ".png"
Public Const DOT_PURPLE As String = ICON_PATH & "dot_purple" & ".png"

'-----------------------------------------------------------------------
' Photo Types
'-----------------------------------------------------------------------
Public Const PHOTO_TYPES_MAIN As String = "Reference,Overview,Feature,Transect,Other"      'String -> basic photo types
Public Const PHOTO_TYPES_OTHER As String = "Animal,Plant,Cultural,Disturbance,Field Work,Scenic,Weather,Other"      'String -> other photo types
Public Const PHOTO_EXT_ALLOWED As String = "jpg,jpeg,png"
Public Const Photo_Path As String = "C:\"
'photo number regex pattern defined in AppEnum

'-----------------------------------------------------------------------
' Upland Components
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' Invasives Components
'-----------------------------------------------------------------------
Public g_AppSurfaces As Scripting.Dictionary            'global application surface names & IDs (for lookups)
Public g_AppQuadratPositions As Scripting.Dictionary    'global application quadrat positions (for lookups)

Public Const QUADRATS_PER_TRANSECT As Integer = 3       'total # of quadrats found on an invasives transect
                                                        'this value assumes quadrat #s are consecutive & begin w/ 1

'-----------------------------------------------------------------------
' Big Rivers Components
'-----------------------------------------------------------------------
Public Const APP_IMAGES_DIR As String = ""
Public Const PARKS = "BLCA,CANY,DINO"
' O - Observer, R - Recorder, DE - DataEntry, V - DataVerify, C - DataCertify
Public Const RECORD_ACTIONS As String = "O,R,DE,V,C"
' O - Observer, R - Recorder, DE - DataEntry, V - DataVerify
' PD - PhotoDownload, P - Photographer, C - DataCertify
Public Const CONTACT_ROLES As String = "O,R,DE,V,C,P,PD"  'add P, PD to db?

Public Const LOCATION_TYPES As String = "F,T,P"     'F=feature, T=transects, P=point

Public Const LINE_DIST_SOURCES As String = "T,P"    'transect & plot

'Measurement type - initially ALL = SC
'WP-water pin, SC-slope change, U-upland, R-river
Public Const LINE_DIST_TYPES As String = "WP,SC,U,R"

'Height of tagline above ...
'H-headpin @ 0, W-water, G-ground, V-vegetation,  WRS - water @ water pin
'SC: Points where tagline bends or stretches while slope changes
'W-water, G-ground, V-vegetation, R- rock, D-debris
Public Const HEIGHT_TYPES As String = "H,W,G,V,WRS,V,R,D"

'Slope Change Causes ...
'V-vegetation, G-ground, W-water, R-rock, D-debris
Public Const SLOPE_CHANGE_CAUSES As String = "D,G,R,V,W"

'Transect, Feature, Reference or Overview (T, F, R, O - transect, feature, reference, overview/point-to-point),
'Other photos: OA-animal, OC-cultural, OD-disturbance, OF-field work, OP-plants, OS-scenic, OW-weather, OO-other
Public Const PHOTO_TYPES As String = "T,F,O,R,OA,OC,OD,OF,OP,OS,OW,OO"

'Transducer types - A-air, W-water
Public Const TRANSDUCER_TYPES As String = "A,W"

'Timing of actions (BD-before-download, AD-after-download/reinstallation)
Public Const TRANSDUCER_TIMING As String = "BD,AD"

'Plot densities
Public Const PLOT_DENSITIES As String = "1,2,4,8"

'Transect numbers --> BLCA & CANY, range 1-8, DINO has no transects
Public Const TRANSECT_NUMBERS As String = "1,2,3,4,5,6,7,8"

'Veg walk collection types --> Site or Feature to handle prior non-site data (S or F)
Public Const COLLECTION_TYPES As String = "S,F"

'Veg plot cover types --> WCC = woody canopy cover (BLCA & CANY)
'                         URC - understory rooted cover (BLCA & CANY),
'                         ARS - all rooted species (DINO)
Public Const COVER_TYPES As String = "WCC,URC,ARS"

'Veg unknowns
Public Const PLANT_TYPES As String = "herb,shrub,tree,grass,sedge,other"  'TEXT(50) --> TEXT(15)
Public Const LEAF_TYPES As String = "compound/simple, arrangement" 'TEXT(50) --> TEXT(25)
Public Const FORB_GRASS_TYPES As String = "Annual,Perennial" 'TEXT(10)
Public Const PERENNIAL_GRASS_TYPES As String = "Bunchgrass, Rhizomatous" 'TEXT(15)
'Salient feature TEXT(255)
'Leaf margin TEXT(50)
'Other leaf characteristics:  pubescence, sap, stipules TEXT(50)
'Stem characteristics: shape, pubescence, bud TEXT(50)
'Flower characteristics: color location floral formula TEXT(50)
'General and microhabitat characteristics TEXT(50)
'Perennial grass type: Bunchgrass or Rhizomatous TEXT(15)
'Collection method TEXT(50)