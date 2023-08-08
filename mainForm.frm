VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainForm 
   Caption         =   "EP-Quick Freeware"
   ClientHeight    =   10545
   ClientLeft      =   705
   ClientTop       =   1320
   ClientWidth     =   16365
   Icon            =   "mainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   16365
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar toolBarMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16365
      _ExtentX        =   28866
      _ExtentY        =   1111
      ButtonWidth     =   1376
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "imgLstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Make IDF"
            Key             =   "createIDF"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grdMain 
      Height          =   9735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9735
      _cx             =   17171
      _cy             =   17171
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"mainForm.frx":08CA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   1
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.PictureBox pctMain 
      Height          =   7575
      Left            =   10080
      ScaleHeight     =   7515
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin MSComDlg.CommonDialog fsDialog 
         Left            =   600
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imgLstToolbar 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mainForm.frx":0932
               Key             =   "new"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mainForm.frx":0E74
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mainForm.frx":13B6
               Key             =   "save"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mainForm.frx":18F8
               Key             =   "print"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mainForm.frx":1E3A
               Key             =   "makeIDF"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New.."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCreateIDF 
         Caption         =   "Make &EnergyPlus IDF File"
      End
      Begin VB.Menu mnuFileDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExt 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpLicense 
         Caption         =   "End Use &License Agreement.."
      End
      Begin VB.Menu mnuHelpRegistration 
         Caption         =   "&Registration..."
      End
      Begin VB.Menu mnuHelpDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About EP-Quick..."
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' IN 2.1 No changes (I think)
' IN 2.2 PEOPLE, LIGHTS, ELECTRIC EQUIPMENT, GAS EQUIPMENT, HOT WATER EQUIPMENT, STEAM EQUIPMENT, OTHER EQUIPMENT, INFILTRATION
' IN 3.0 Names of all objects.


'------------------------------------------------------------------------
'
' EP-Quick (code named EPoxy)
'
' The template based user interface for EnergyPlus
'
' (c) 2003-2009 by Jason Glazer. All rights reserved.
'
' EPoxy is a simple program that allows users to choose from a bunch of
' templates that all a partial IDF file to be generated quickly.  Instead
' of specifying the location of all of the coordinates in a buildings
' geometry, the user specifies a series of parameters that specify overall
' dimensions such as number of floors, width, and length.  The templates
' are located in the templates subdirectory and contain the following inputs:
'
' input, description, variable, default, min, max, units
' rule, expression
' corner, cornername, x-expression, y-expression
' roofcorner,cornername, x-expression, y-expression
' extwall, cornername1, cornername2
' intwall, cornername1, cornername2
' roof, cornername1, cornername2, cornername3, cornername4, ...
' zone, cornername1, cornername2, cornername3, cornername4, ...
'
' A character ! may be used at the beginning of comment lines
' Input need to listed first
' Corners need to be listed prior to extwall, intwall, and zone
' For roof and zone they should be entered in clockwise rotation if seen from
' above the building

'------------------------------------------------------------------------
' Development Notes
'
' DONE
'    file clear/new
'    exit
'    save/open
'    convert to SI for everything
'    Help about
'    Generate IDF
'    IP/SI setting needs to be saved
'    WindowGas not showing up
'    Include DXF report for building
'    Include output variables
'    Simple Templates
'    Add door constructions
'    Fix rotation coord computation
'    Beta time limited version
'    Building with roof and attic
'    Fix insulation conversion factor
'    Help About EPoxy should be EP-Quick
'    Inform user that IDF is in same directory as EPQ
'    Defaults in metric mode not converted
' DONE FOR SECOND PUBLIC BETA
'    Defaults for SI units still IP values (Dru Aug 24 2004)
'    Multiple doors have the same name (Raustad Aug 4 2004)
'    Windows multiplier with solar dist warning (Noe Aug 8 2004)
'    Added roof corners and roof pieces to all templates fixing the following three errors
'    but still get warning message (Sep 10 2004)
'      Roof peak height in rectangular two zone ignored (Noe Aug 8 2004)
'      Roof peak height rectangular one zone ignored (Noe Aug 8 2004)
'      Roof peak height in Rectangle with Perimeters Corners and Core ignored (Noe Aug 8 2004)
'    Fix version number written to 1.2.1
'    ABUPS on partial year gives warning
'    Box9noroof and box9roof core zone not enough surfaces
'    Roof peak height in Rectangle with Perimeter and Core warning message (Noe Aug 8 2004)
'      No longer warning message
'    Web site/gard involvement
'    Public beta
'
' DONE FOR RELEASE
'    More templates
'    Polygon templates
'    Warn if polygon template selected
'    Set price $79 normal, $29 student, $0 for EnergyPlus contributor
'    Allow for short run period or entire year
'    Registration/limited functionality
'    Write help
'    Define multiple windows instead of using multiplier so that shading can be full
'    Fixed bug when opening file and no drawing found
'    Make PDF of help and help file
'    Create and link help file
'
' BUGS FOUND IN ENERGYPLUS TEAM BETA
'    Internal mass - input - gyp with floor area?
'
' BUGS FOUND IN PUBLIC BETA AS OF SEPT 3 2004
'    Box2roof shows floor multiplier and interzone surface problems with roof
'    TriRoof warning about number of surfaces
'    Box1roof shows warning about less than 6 surfaces in attic
'    MAKE SURE ANNUAL RUNS ARE ACTIVE
'
' TO DO FOR RELEASE
'    Run test protocol
'    Update web site for purchase
'    Add template list and description to web site
'
' TEST PROTOCOL
'    Run the following cases:
'       All templates with defaults
'       All templates with tilt roof
'       At least three templates with design days
'       At least one template with windows and doors (several different)
'       At least one template with three different internal gains
'       At least one template with top floor different
'       At least one template with bottom floor different
'       At least one template with basement
'       At least one template with SI
'       At least one template with building rotation
'       At least one template with much larger size
'   The three templates shall be:
'       Single zone box
'       Multizone non-polygon
'       Polygon
'
' Suggestions from beta test users for future versions
'    Add purchased air or system/plant
'    Save current file location in registry
'    Add in internal gains the sqft for each type and percent of building
'    Cancel button on "new" frame
'    Allow input as window to wall ratio
'    Why World Coordinate System
'    Not clear where are defaults
'    Internal gains are confusing
'    Smarting defaulting so that geometry of subsurfaces is consistant with parent surface (Raustad)
'    PDF data sheets of entries needed (krenzel Aug 12 2004)
'    Full screen/sizable screen (Crawley Oct 2004)
'    Copy/paste
'    Print
'    Overhang on windows
'    Setback on windows
'    HVAC inputs
'    Construction editor
'    Style/schedule editor
'    Use multiplier objects
'    Fix pitched roof view factor to ground (???)
'------------------------------------------------------------------------

Option Explicit

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Public programVersion As String

'------------------------------------------------------------------------
' Variables that are for saving and loading files
'------------------------------------------------------------------------
Dim curFileName As String
Dim curFilePath As String
Dim curFileNameWithPath As String
Dim fileChangedSinceSave As Boolean
Dim fileIsUntitled As Boolean
Dim appPathSlash As String
Dim idfFileHandle As Integer
Dim warnUserWhereisIDF As Boolean

'------------------------------------------------------------------------
' Types and variable declarations for plan file descriptions
'------------------------------------------------------------------------
Const maxNumFloorPlan = 4

Private Type pUserInputType
  curVal As Single
  Description As String
  variable As String
  default As Single
  min As Single
  max As Single
End Type
Const maxNumPUserInput = 30
Dim pUserInput(maxNumPUserInput) As pUserInputType
Dim numPUserInput As Integer

Const maxNumPRules = 30
Private Type pRuleType
  isGood As Boolean
  Expression As String
  UserInCnt As Integer
  UserIn(maxNumPUserInput) As Integer
End Type
Dim pRule(maxNumPRules) As pRuleType
Dim numPRule As Integer

Private Type pCornerType
  name As String
  xexpression As String
  yexpression As String
  x As Single  'computed - not part of PLN file input
  y As Single  'computed - not part of PLN file input
  'used in setting the order of the user input for the expressions
  xUserInCnt As Integer
  xUserIn(maxNumPUserInput) As Integer
  yUserInCnt As Integer
  yUserIn(maxNumPUserInput) As Integer
  xTrans As Single    'after transformation due to building rotation
  yTrans As Single    'after transformation due to building rotation
  xTransSI As Single       'convert to SI units (includes transformation)
  yTransSI As Single       'convert to SI units (includes transformation)
End Type
Const maxNumPCorners = 100
Dim pCorner(maxNumPCorners) As pCornerType
Dim numPCorner As Integer
Const linesCorner = 6
Const maxNumPRoofCorners = 25
Dim pRoofCorner(maxNumPRoofCorners) As pCornerType
Dim numPRoofCorner As Integer
Const linesRoofCorner = 6

Private Type pWallType
  nm(maxNumFloorPlan) As String 'concatenation of both corner names
  startCorner As Integer
  endCorner As Integer
  length As Single  'computed - not part of PLN file input
  lengthSI As Single 'SI value of length
  area(maxNumFloorPlan) As Single    'computed - not part of PLN file input
  areaSI(maxNumFloorPlan) As Single  'SI value of area
  cons(maxNumFloorPlan) As Integer   'user choice (exterior only) - not part of PLN file input
  insul(maxNumFloorPlan) As Integer   'user choice (exterior only) - not part of PLN file input
  perGlaz(maxNumFloorPlan) As Single 'computed (exterior only) - not part of PLN file input
  glazArea(maxNumFloorPlan) As Single 'computed - used internally only
  zone1 As Integer 'computed - zone that this wall is associated with
  zone2 As Integer 'computed - second zone that this interior wall is associated with
  consInsul(maxNumFloorPlan) As Integer 'used for constructions with insulation
End Type
Const maxNumPWalls = 100
Dim pExtWall(maxNumPWalls) As pWallType
Dim numPExtWall As Integer
Dim pIntWall(maxNumPWalls) As pWallType
Dim numPIntWall As Integer
Const linesExtWall = 5
Const linesIntWall = 2

Const maxNumZoneCrnrs = 20
Private Type pZoneType
  nm(maxNumFloorPlan) As String 'concatenated list of corner names
  numZoneCrnrs As Integer
  crnrs(maxNumZoneCrnrs) As Integer
  area As Single  'computed - not part of PLN file input
  areaSI As Single 'SI value of area (metric)
  style(maxNumFloorPlan) As Integer 'user choice - not part of PLN file input
End Type
Const maxNumPZones = 100
Dim pZone(maxNumPZones) As pZoneType
Dim numPZone As Integer
Const linesZone = 2

Const maxNumRoofCrnrs = 20
Private Type pRoofType
  nm As String 'concatenated list of corner names
  numRoofCrnrs As Integer
  crnrs(maxNumRoofCrnrs) As Integer
  area As Single  'computed - not part of PLN file input
  areaSI As Single 'SI value of area (metric)
End Type
Const maxNumPRoof = 100
Dim pRoof(maxNumPRoof) As pRoofType
Dim numPRoof As Integer
Const linesRoof = 1

'------------------------------------------------------------------------
' Types and variable declarations for each input group
' Constants for the number of lines
'------------------------------------------------------------------------

Const linesBuilding = 14
Private Type iBuildingType
  roofCons As Integer
  roofInsul As Integer
  intWallCons As Integer
  floorCons As Integer
  botFloorCons As Integer
  botFloorInsul As Integer
  northAngle As Single
  perGlaz As Single  'computed
  floorArea As Single 'computed
  wallArea As Single  'computed
  glazArea As Single  'computed
  height As Single    'computed
  roofPkHt As Single
  roofPkHtSI As Single 'metric value of roofPkHt
  duration As Integer
  planName As String
  epVersion As Integer
End Type
Dim iBuilding As iBuildingType
Dim durationAnnual As Integer
Dim durationDesign As Integer

Const linesDefault = 12
Private Type iDefaultType
  flr2flr As Single
  flr2flrSI As Single 'metric (SI) value of flr2flr
  style As Integer
  extWallCons As Integer
  extWallInsul As Integer
  windCons As Integer
  windWidth As Single
  windWidthSI As Single 'metric (SI) value of windWidth
  windHeight As Single
  windHeightSI As Single 'metric (SI) value of windHeight
  windOvrhng As Single
  windOvrhngSI As Single 'metric (SI) value of windOvrhng
  windSetbck As Single
  windSetbckSI As Single 'metric (SI) value of windSetbck
  doorCons As Integer
  doorWidth As Single
  doorWidthSI As Single 'metric (SI) value of doorWidth
  doorHeight As Single
  doorHeightSI As Single 'metric (SI) value of doorHeight
End Type
Dim iDefault As iDefaultType

' now called internal gains
Const linesStyle = 10
Private Type iStyleType
  nm As String
  isUsed As Boolean
  peopDensUse As Single
  peopDensUseSI As Single 'metric (SI) value of peopDens
  peopDensNonUse As Single
  peopDensNonUseSI As Single 'metric (SI) value of peopDens
  liteDensUse As Single
  liteDensUseSI As Single 'metric (SI) value of liteDens
  liteDensNonUse As Single
  liteDensNonUseSI As Single 'metric (SI) value of liteDens
  eqpDensUse As Single
  eqpDensUseSI As Single 'metric (SI) value of eqpDens
  eqpDensNonUse As Single
  eqpDensNonUseSI As Single 'metric (SI) value of eqpDens
  weekdayTimeRange As Integer
  saturdayTimeRange As Integer
  sundayTimeRange As Integer
  furnDens As Single
  furnDensSI As Single 'metric (SI) value of furnDens
End Type
Const numIStyle = 10
Dim iStyle(numIStyle) As iStyleType

' Active indicates if the user has specified those floors
' as different.  The floors are:
'       1 = basement
'       2 = lower
'       3 = middle (or all)
'       4 = top
Const basementFloor = 1
Const lowerFloor = 2
Const middleFloor = 3
Const topFloor = 4
Const linesFloorplan = 3
Private Type iFloorPlanType
  nm As String
  active As Boolean
  flr2flr As Single
  flr2flrSI As Single 'metric (SI) value of flr2flr
  numFlr As Single
  flrArea As Single 'computed
  heightOfFloor As Single 'computed
  heightOfFloorSI As Single 'metric (SI) value of heightOfFloor
End Type
Dim iFloorPlan(maxNumFloorPlan) As iFloorPlanType
Dim numFloorPlans As Integer

Const windowsPerWall = 3
Const doorsPerWall = 2

Const linesWindow = 4
Private Type iWindowType
  nm As String
  cons As Integer
  width As Single
  widthSI As Single 'metric (SI) value of width
  height As Single
  heightSI As Single 'metric (SI) value of height
  count As Single
  ' the following are not currently implemented so should be ignored
  ovrhng As Single
  ovrhngSI As Single 'metric (SI) value of ovrhng
  setbck As Single
  setbckSI As Single 'metric (SI) value of setbck
End Type
Dim iWindow(maxNumFloorPlan, maxNumPWalls, windowsPerWall) As iWindowType

Const linesDoor = 4
Private Type iDoorType
  nm As String
  cons As Integer
  width As Single
  widthSI As Single 'metric (SI) value of width
  height As Single
  heightSI As Single 'metric (SI) value of height
  count As Single
End Type
Dim iDoor(maxNumFloorPlan, maxNumPWalls, doorsPerWall) As iDoorType


Private Type materialType
  nm As String         'name of the material (HOF reference)
  desc As String       'description of material
  isUsed As Boolean    'flag if the material is used by some construction
  rough As Integer     '1=smooth
  thick As Single      'thickness - m
  conduct As Single    'conductivity - w/m-k
  dens As Single       'density - kg/m3
  spheat As Single     'specific heat - J/kg-K
  emit As Single       'thermal emittance
  solAbs As Single     'solar absorptance
  visAbs As Single     'visible absorptance
End Type
Const numMaterials = 50
Dim MATERIAL(numMaterials) As materialType
Dim gypForInterior As Integer         'used to flag the gypsum which should be used for interior walls
Dim concForInterior As Integer        'used to flag the concrete which should be used for interior walls

Private Type constLayerType
  nm As String         'root name used in particular construction
  isUsed As Boolean    'flag if the construction is used
  matCount As Integer  'number of materials in the layer
  matInd(20) As Integer 'pointer to the material array for each layer
End Type
Const numConstLayer = 100
Dim constLayer(numConstLayer) As constLayerType
Const constLayerInsul = -1
Const constLayerAirGap = -2

Private Type insulationType
  nm As String   'INSUL R-x
  rValue As Integer
  isUsed As Boolean
End Type
Const numInsulation = 30
Dim insulation(numInsulation) As insulationType

Private Type constInsulComboType
  nm As String
  insulationPt As Integer
  constLayerPt As Integer
End Type
Const maxNumConstInsul = 300
Dim constInsulCombo(maxNumConstInsul) As constInsulComboType
Dim numConstInsulCombo As Integer
Dim roofConsInsul As Integer

Private Type windowGlassGasType
  nm As String
  prop As String  'the entire object
  isUsed As Boolean
End Type
Const numWindowGlassGas = 78
Dim windowGlassGas(numWindowGlassGas) As windowGlassGasType

Private Type windowLayersType
  isUsed As Boolean
  nm As String
  layerCount As Integer
  layerName(7) As String
End Type
Const numWindowLayers = 220
Dim windowLayers(0 To numWindowLayers) As windowLayersType
'------------------------------------------------------------------------
' Define arrays for lists of choice parameters
'------------------------------------------------------------------------
Const maxNumListOfChoices = 1000
Dim listOfChoices(maxNumListOfChoices) As String

Private Type kindOfListType
  firstChoice As Integer
  lastChoice As Integer
  builtString As String
End Type

Const useDefault = 0  'points to general option in lists called Use Default
Const useNumericDefault = -999 'flag value for use default for numeric fields
Const indicateDefault = "Use Default"

Const maxNumKindOfList = 100
Dim kindOfList(maxNumKindOfList) As kindOfListType

Dim listWallConstruction As Integer
Dim defaultWallConstruction As Integer

Dim listRoofConstruction As Integer
Dim defaultRoofConstruction As Integer

Dim listFloorConstruction As Integer
Dim defaultFloorConstruction As Integer

Dim listInsulation As Integer
Dim defaultInsulation As Integer

Dim listStyle As Integer
Dim defaultStyle As Integer

Dim listWindow As Integer
Dim defaultWindow As Integer

Dim listDoor As Integer
Dim defaultDoor As Integer

Dim listSchedule As Integer
Dim defaultSchedule As Integer

Dim listIntWallCons As Integer
Dim defaultIntWallCons As Integer

Dim listTimeRange As Integer
Dim defaultTimeRange As Integer

Dim listDuration As Integer
Dim defaultDuration As Integer

Dim listEPversion As Integer
Dim defaultEPversion As Integer
Dim epVersion121 As Integer
Dim epVersion122 As Integer
Dim epVersion123 As Integer
Dim epVersion130 As Integer
Dim epVersion140 As Integer
Dim epVersion200 As Integer
Dim epVersion210 As Integer
Dim epVersion220 As Integer
Dim epVersion300 As Integer
Dim epVersion310 As Integer
Dim epVersion400 As Integer


'------------------------------------------------------------------------
' Define the heirarchy of input displayed
'------------------------------------------------------------------------
Const kindNumeric = 1
Const kindShowOnlyNumeric = 2
Const kindShowOnlyString = 2
Const kindList = 3
Const kindGroupTitle = 4

Const prrPopulate = 1
Const prrRefresh = 2
Const prrRead = 3
Const prrRepopulate = 4

'------------------------------------------------------------------------
' Type for information needed to display and edit a particular row
' of the grid. This array is dynamically sized with the size of the
' grid.
'------------------------------------------------------------------------
Private Type rowDataType
  kindOfRow As Integer
  defaultVal As Single
  minVal As Single
  maxVal As Single
  defaultMinMaxString As String
  listOfOptionsString As String
End Type
Dim rowData() As rowDataType


'------------------------------------------------------------------------
' Compute objects
'------------------------------------------------------------------------
Dim computeCornerX(maxNumPCorners) As New clsMathParser
Dim computeCornerY(maxNumPCorners) As New clsMathParser
Dim computeRoofCornerX(maxNumPRoofCorners) As New clsMathParser
Dim computeRoofCornerY(maxNumPRoofCorners) As New clsMathParser
Dim computeRules(maxNumPRules) As New clsMathParser


'------------------------------------------------------------------------
' When the program first starts up - the sequence of initialization
'------------------------------------------------------------------------
Private Sub Form_Load()
  programVersion = 1.6
  'curDate = Date
  'endDate = DateSerial(2004, 11, 15)
  warnUserWhereisIDF = True
  mnuHelpRegistration.Visible = False
  mnuHelpDiv2.Visible = False
  App.HelpFile = App.Path & "\ep-quick.HLP"
  'Call showSplashScreen
  appPathSlash = App.Path
  If Right(appPathSlash, 1) <> "\" Then appPathSlash = appPathSlash & "\"
  fileIsUntitled = True
  grdMain.Rows = 1
  'check if beta period is over
  'If DateDiff("d", curDate, endDate) < 0 Then
  '  MsgBox "Sorry the beta period has ended", vbInformation, "EP-Quick"
  '  Unload Me
  '  End
  'End If
  'initialize
  Call initializeChoiceList
  Call initializeStyle
  Call initializeMaterials
  'Call doNewFile
  Call updateWindowTitleBar
End Sub

'------------------------------------------------------------------------
' display the splashscreen
'------------------------------------------------------------------------
Sub showSplashScreen()
frmSplash.Show vbModal
End Sub

'------------------------------------------------------------------------
' This event is called when the X box is clicked on in the application
' should be treated just the same as a file-exit command.
'------------------------------------------------------------------------
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
  Call doExitProgram
End Sub

'------------------------------------------------------------------------
' Set up the type of edit prior
'------------------------------------------------------------------------
Private Sub grdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col < 2 Then
  cancel = True 'for the first two columns do not allow editing
Else
  grdMain.ComboList = ""  'reset to blank value
  Select Case rowData(Row).kindOfRow
    Case kindShowOnlyNumeric
      cancel = True 'for read only do not allow editing
    Case kindShowOnlyString
      cancel = True 'for read only do not allow editing
    Case kindGroupTitle
      cancel = True 'for title do not allow editing
    Case kindList
      grdMain.ComboList = rowData(Row).listOfOptionsString
    Case kindNumeric
      grdMain.ComboList = "|" & grdMain.TextMatrix(Row, Col) & vbTab & "current" & rowData(Row).defaultMinMaxString
  End Select
End If
End Sub

'------------------------------------------------------------------------
' Confirm what the user has entered into the cell just prior to it being
' committed as the contents of the cell.
'------------------------------------------------------------------------
Private Sub grdMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
Dim numEditText As Single
Select Case rowData(Row).kindOfRow
  Case kindShowOnlyNumeric
    cancel = False
  Case kindShowOnlyString
    cancel = False
  Case kindGroupTitle
    cancel = False
  Case kindList
    cancel = False
  Case kindNumeric
    If grdMain.EditText = indicateDefault Then
      cancel = False
    ElseIf Not IsNumeric(grdMain.EditText) Then
      cancel = True
    Else  'is numeric
      numEditText = Val(grdMain.EditText)
      If numEditText < rowData(Row).minVal Then
        cancel = True
      ElseIf numEditText > rowData(Row).maxVal Then
        cancel = True
      Else
        cancel = False
      End If
    End If
End Select
End Sub

'------------------------------------------------------------------------
' After a value has been changed, recompute everything
'------------------------------------------------------------------------
Private Sub grdMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Select Case rowData(Row).kindOfRow
  Case kindNumeric
    Call populateRefreshReadGrid(prrRead)
    Call recompute
    fileChangedSinceSave = True
    Call updateWindowTitleBar
    Call populateRefreshReadGrid(prrRefresh)
  Case kindList
    Call populateRefreshReadGrid(prrRead)
    fileChangedSinceSave = True
    Call updateWindowTitleBar
End Select
End Sub



'------------------------------------------------------------------------
' Handle the toolbar button events
'------------------------------------------------------------------------
Private Sub toolBarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case UCase(Button.Key)
    Case "NEW"
      If grdMain.EditWindow <> 0 Then
        grdMain.Select grdMain.Row, grdMain.Col
      End If
      Call clearAll
      Call doNewFile
    Case "OPEN"
      Call doOpenFile
    Case "SAVE"
      Call doSaveFile(0)
    Case "CREATEIDF"
      Call doCreateIDF
  End Select
End Sub

'------------------------------------------------------------------------
' Handle the File Menu events
'------------------------------------------------------------------------
Private Sub mnuFileNew_Click()
  Call doNewFile
End Sub
Private Sub mnuFileOpen_Click()
  Call doOpenFile
End Sub
Private Sub mnuFileSave_Click()
  Call doSaveFile(0)
End Sub
Private Sub mnuFileSaveAs_Click()
  Call doSaveFile(1)
End Sub
Private Sub mnuFileExt_Click()
  Call doExitProgram
End Sub
Private Sub mnuFileCreateIDF_Click()
  Call doCreateIDF
End Sub

'------------------------------------------------------------------------
' Handle the Help Menu events
'------------------------------------------------------------------------
Private Sub mnuHelpAbout_Click()
frmSplash.Show vbModal
End Sub
Private Sub mnuHelpContents_Click()
Dim nRet As Integer
If Len(App.HelpFile) = 0 Then
  MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
Else
  On Error Resume Next
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
  If Err Then
    MsgBox Err.Description
  End If
End If
End Sub
Private Sub mnuHelpLicense_Click()
frmEULA.Show vbModal
End Sub

'------------------------------------------------------------------------
' Show New Dialog Box and Initialize New File
'------------------------------------------------------------------------
Sub doNewFile()
  Dim yesOrNo As Integer
  Dim ln As Long
  Dim templateName As String
  If fileChangedSinceSave Then
    yesOrNo = MsgBox("The file has changed since the last time it was saved.  Do you want to save the file before opening a new template?", vbYesNo, "File Changed Warning")
    If yesOrNo = vbYes Then
      Call doSaveFile(0)
    End If
  End If
  newPlan.Show vbModal
  Call readPlanFile
  Call loadExpressions
  Call setActiveFloors
  Call formZoneWallRoofNames
  Call initializeStyle
  Call buildChoiceListString
  Call populateRefreshReadGrid(prrPopulate)
  On Error Resume Next
  templateName = App.Path & "\planTemplates\" & newPlanInfo.templateName & ".wmf"
  ln = FileLen(templateName)
  If Err.Number = 0 Then
    pctMain.Picture = LoadPicture(templateName)
  Else
    pctMain.Picture = LoadPicture(App.Path & "\planTemplates\NoDrawingFound.wmf")
  End If
  Call recompute
  Call populateRefreshReadGrid(prrRefresh)
  fileChangedSinceSave = False
  fileIsUntitled = True
  curFileNameWithPath = "c:\untitled.epq"
  curFileName = extractFileNameNoExt(curFileNameWithPath)
  curFilePath = extractPath(curFileNameWithPath)
  Call updateWindowTitleBar
End Sub

'------------------------------------------------------------------------
' Open an existing file
'------------------------------------------------------------------------
Sub doOpenFile()
Dim wasItCancelled As Boolean
Dim ln As Long
Dim templateName As String
Call openLocalFileDialog(wasItCancelled)
If wasItCancelled Then Exit Sub
Call clearAll
Call setWindowsDoorsToDefault
Call initializeStyle
Call readActive
Call loadExpressions
Call formZoneWallRoofNames
Call buildChoiceListString
Call populateRefreshReadGrid(prrRepopulate)
On Error Resume Next
templateName = App.Path & "\planTemplates\" & iBuilding.planName & ".wmf"
ln = FileLen(templateName)
If Err.Number = 0 Then
  pctMain.Picture = LoadPicture(templateName)
Else
  pctMain.Picture = LoadPicture(App.Path & "\planTemplates\NoDrawingFound.wmf")
End If
Call recompute
Call populateRefreshReadGrid(prrRefresh)
fileChangedSinceSave = False
fileIsUntitled = False
Call updateWindowTitleBar
End Sub

'------------------------------------------------------------------------
' Save the file
'  If typeOfSave = 0 then normal save
'                = 1 then save as
'------------------------------------------------------------------------
Sub doSaveFile(typeOfSave As Integer)
Dim wasItCancelled As Boolean
If fileIsUntitled Or typeOfSave = 1 Then
  Call useFileSaveDialog(wasItCancelled)
  If wasItCancelled Then Exit Sub
End If
Call saveActive
fileChangedSinceSave = False
fileIsUntitled = False
Call updateWindowTitleBar
End Sub

'------------------------------------------------------------------------
' Check if files are saved,
'------------------------------------------------------------------------
Sub doExitProgram()
  Dim yesOrNo As Integer
  If fileChangedSinceSave Then
    yesOrNo = MsgBox("The file has changed since the last time it was saved.  Do you want to save the file before exiting?", vbYesNo, "File Changed Warning")
    If yesOrNo = vbYes Then
      Call doSaveFile(0)
    End If
  End If
  End
End Sub


'------------------------------------------------------------------------
' read the plan file/string and initialize the arrays
' either way the saveplan.fil file is created that mirrors the template
'------------------------------------------------------------------------
Sub readPlanFile()
Dim inFn As Integer
Dim lineFromFile As String
Dim partsOfLine() As String
Dim numOfParts As Integer
Dim i As Integer
On Error Resume Next
inFn = FreeFile
' if opening up a new template then find it and
' check if it is present
If newPlanInfo.templateName = "" Then Exit Sub
iBuilding.planName = newPlanInfo.templateName
Open appPathSlash & "planTemplates\" & newPlanInfo.templateName & ".pln" For Input As inFn
'loop through the open file or the template file and parse the template related lines
Do While Not EOF(inFn)
  Line Input #inFn, lineFromFile
  'skip lines with comment character !
  If Left(Trim(lineFromFile), 1) <> "!" Then
    'separate the line read into pieces
    partsOfLine = Split(lineFromFile, ",", -1)
    numOfParts = UBound(partsOfLine)
    If numOfParts >= 1 Then
      Select Case LCase(partsOfLine(0))
        Case "input"
          numPUserInput = numPUserInput + 1
          If numPUserInput > maxNumPUserInput Then
            MsgBox "Too many input fields defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pUserInput(numPUserInput).Description = partsOfLine(1)
          pUserInput(numPUserInput).variable = partsOfLine(2)
          pUserInput(numPUserInput).default = CSng(partsOfLine(3))
          pUserInput(numPUserInput).min = CSng(partsOfLine(4))
          pUserInput(numPUserInput).max = CSng(partsOfLine(5))
        Case "rule"
          numPRule = numPRule + 1
          If numPRule > maxNumPRules Then
            MsgBox "Too many rules defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pRule(numPRule).Expression = partsOfLine(1)
        Case "corner"
          numPCorner = numPCorner + 1
          If numPCorner > maxNumPCorners Then
            MsgBox "To many corners defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pCorner(numPCorner).name = partsOfLine(1) 'name of corner
          pCorner(numPCorner).xexpression = partsOfLine(2)
          pCorner(numPCorner).yexpression = partsOfLine(3)
        Case "roofcorner"
          numPRoofCorner = numPRoofCorner + 1
          If numPRoofCorner > maxNumPRoofCorners Then
            MsgBox "To many roof corners defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pRoofCorner(numPRoofCorner).name = partsOfLine(1) 'name of roof corner
          pRoofCorner(numPRoofCorner).xexpression = partsOfLine(2)
          pRoofCorner(numPRoofCorner).yexpression = partsOfLine(3)
        Case "extwall"
          numPExtWall = numPExtWall + 1
          If numPExtWall > maxNumPWalls Then
            MsgBox "To many extwalls defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pExtWall(numPExtWall).startCorner = idCorner(partsOfLine(1))
          pExtWall(numPExtWall).endCorner = idCorner(partsOfLine(2))
          If pExtWall(numPExtWall).startCorner < 0 Or pExtWall(numPExtWall).endCorner < 0 Then
            MsgBox "Could not find corner named in extWall", vbCritical, "Reading Template File"
            Exit Sub
          End If
        Case "intwall"
          numPIntWall = numPIntWall + 1
          If numPIntWall > maxNumPWalls Then
            MsgBox "To many intwalls defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pIntWall(numPIntWall).startCorner = idCorner(partsOfLine(1))
          pIntWall(numPIntWall).endCorner = idCorner(partsOfLine(2))
          If pIntWall(numPIntWall).startCorner < 0 Or pIntWall(numPIntWall).endCorner < 0 Then
            MsgBox "Could not find corner named in intWall", vbCritical, "Reading Template File"
            Exit Sub
          End If
        Case "zone"
          numPZone = numPZone + 1
          If numPZone > maxNumPZones Then
            MsgBox "Too many zones defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pZone(numPZone).numZoneCrnrs = numOfParts
          For i = 1 To numOfParts
            pZone(numPZone).crnrs(i) = idCorner(partsOfLine(i))
          Next i
        Case "roof"
          numPRoof = numPRoof + 1
          If numPRoof > maxNumPRoof Then
            MsgBox "Too many roof pieces defined in template", vbCritical, "Reading Template File"
            Exit Sub
          End If
          pRoof(numPRoof).numRoofCrnrs = numOfParts
          For i = 1 To numOfParts
            pRoof(numPRoof).crnrs(i) = idCorner(partsOfLine(i))
          Next i
      End Select
    Else
      MsgBox "Line of template file cannot be parsed:" & vbCrLf & vbCrLf & lineFromFile & vbCrLf & partsOfLine(0), vbInformation, "Reading Template File"
    End If
  End If
Loop
Close inFn
End Sub

'------------------------------------------------------------------------
' Take name of corner and return index to
' the named corner
'    corners are associated with positive numbers
'    roof corners return negative numbers
'    no match found are zero numbers
'------------------------------------------------------------------------
Function idCorner(nameOfCorner As String) As Integer
Dim noc As String
Dim foundCorner As Integer
Dim i As Integer
noc = LCase(Trim(nameOfCorner))
foundCorner = 0
For i = 1 To numPCorner
  If LCase(pCorner(i).name) = noc Then
    foundCorner = i
    Exit For
  End If
Next i
If foundCorner = 0 Then
  For i = 1 To numPRoofCorner
    If LCase(pRoofCorner(i).name) = noc Then
      foundCorner = -i
      Exit For
    End If
  Next i
End If
If foundCorner = 0 Then
  MsgBox "Corner named: " & nameOfCorner & " not found", vbExclamation, "Parsing error."
End If
idCorner = foundCorner
End Function

'------------------------------------------------------------------------
' Populate, Refresh and Read the control grid
'
' A single routine doing the populating, refreshing and reading of the
' grid will make it easier to maintain since all operations are
' basically done together.
'------------------------------------------------------------------------
Sub populateRefreshReadGrid(popRefRead As Integer)
Dim numLines As Integer
Dim curRowNum As Integer
Dim buildingGroupRow As Integer
Dim defaultGroupRow As Integer
Dim styleGroupRow As Integer
Dim floorplanGroupRow As Integer
Dim cornerGroupRow As Integer
Dim roofCornerGroupRow As Integer
Dim i As Integer, j As Integer, k As Integer
If popRefRead = prrPopulate Or popRefRead = prrRepopulate Then
  ' count the number of lines in the grid
  numLines = 0
  numLines = numLines + numPUserInput + 1
  numLines = numLines + linesBuilding + 1
  numLines = numLines + linesDefault + 1
  numLines = numLines + (1 + linesCorner) * numPCorner + 1
  numLines = numLines + (1 + linesCorner) * numPRoofCorner + 1
  numLines = numLines + (1 + linesStyle) * numIStyle + 1
  numLines = numLines + (1 + linesFloorplan) * numFloorPlans + 1
  numLines = numLines + ((1 + linesZone) * numPZone + 1) * numFloorPlans
  numLines = numLines + ((1 + linesIntWall) * numPIntWall + 1) * numFloorPlans
  numLines = numLines + ((1 + linesExtWall) * numPExtWall + 1) * numFloorPlans
  numLines = numLines + ((1 + linesWindow) * numPExtWall * windowsPerWall + 1) * numFloorPlans
  numLines = numLines + ((1 + linesDoor) * numPExtWall * doorsPerWall + 1) * numFloorPlans
  ' Size the grid so that it contains all the rows based on the estimate
  grdMain.Rows = numLines
  ' size the dynamic rowData array that contains information about how to edit
  ' a row to the same size as the grid
  ReDim rowData(numLines)
End If
curRowNum = 0
' geometry (pUserInput)
Call prrGroupTitle(popRefRead, curRowNum, 1, "Overall Geometry")
For i = 1 To numPUserInput
  Call prrNumeric(popRefRead, curRowNum, pUserInput(i).curVal, pUserInput(i).variable, pUserInput(i).default, pUserInput(i).min, pUserInput(i).max, "ft", "m")
  Debug.Print pUserInput(i).curVal, pUserInput(i).variable, pUserInput(i).default
Next i
' building
Call prrGroupTitle(popRefRead, curRowNum, 1, "Building")
buildingGroupRow = curRowNum
Call prrList(popRefRead, curRowNum, iBuilding.roofCons, "Roof Construction", defaultRoofConstruction, listRoofConstruction)
Call prrList(popRefRead, curRowNum, iBuilding.roofInsul, "Roof Insulation", defaultInsulation, listInsulation)
Call prrList(popRefRead, curRowNum, iBuilding.intWallCons, "Interior Wall Construction", defaultIntWallCons, listIntWallCons)
Call prrList(popRefRead, curRowNum, iBuilding.floorCons, "Floor Construction", defaultFloorConstruction, listFloorConstruction)
Call prrList(popRefRead, curRowNum, iBuilding.botFloorCons, "Bottom Floor Construction", defaultFloorConstruction, listFloorConstruction)
Call prrList(popRefRead, curRowNum, iBuilding.botFloorInsul, "Roof Insulation", defaultInsulation, listInsulation)
Call prrNumeric(popRefRead, curRowNum, iBuilding.northAngle, "North Angle", 0, 0, 359.9, "deg", "deg")
Call prrShowOnlyNumeric(popRefRead, curRowNum, iBuilding.wallArea, "Wall Area", "sqft", "m2")
Call prrShowOnlyNumeric(popRefRead, curRowNum, iBuilding.perGlaz, "Percent Glazing", "%", "%")
Call prrShowOnlyNumeric(popRefRead, curRowNum, iBuilding.floorArea, "Floor Area", "sqft", "m2")
Call prrShowOnlyNumeric(popRefRead, curRowNum, iBuilding.height, "Height", "ft", "m")
Call prrNumeric(popRefRead, curRowNum, iBuilding.roofPkHt, "Roof Peak Height", 0, 0, 40, "ft", "m")
Call prrList(popRefRead, curRowNum, iBuilding.duration, "Duration", defaultDuration, listDuration)
Call prrShowOnlyString(popRefRead, curRowNum, iBuilding.planName, "Template Name")
Call prrList(popRefRead, curRowNum, iBuilding.epVersion, "EnergyPlus Version", defaultEPversion, listEPversion)
' defaults
Call prrGroupTitle(popRefRead, curRowNum, 1, "Defaults")
defaultGroupRow = curRowNum
Call prrNumeric(popRefRead, curRowNum, iDefault.flr2flr, "Floor to Floor Height", 10, 1, 40, "ft", "m")
Call prrList(popRefRead, curRowNum, iDefault.style, "Internal Gains", defaultStyle, listStyle)
Call prrList(popRefRead, curRowNum, iDefault.extWallCons, "Ext Wall Construction", defaultWallConstruction, listWallConstruction)
Call prrList(popRefRead, curRowNum, iDefault.extWallInsul, "Ext Wall Insulation", defaultInsulation, listInsulation)
Call prrList(popRefRead, curRowNum, iDefault.windCons, "Window Type", defaultWindow, listWindow)
Call prrNumeric(popRefRead, curRowNum, iDefault.windWidth, "Window Width", 3, 1, 12, "ft", "m")
Call prrNumeric(popRefRead, curRowNum, iDefault.windHeight, "Window Height", 5, 1, 50, "ft", "m")
Call prrNumeric(popRefRead, curRowNum, iDefault.windOvrhng, "Window Overhang", 0, 0, 20, "ft", "m")
Call prrNumeric(popRefRead, curRowNum, iDefault.windSetbck, "Window Setback", 0, 0, 10, "ft", "m")
Call prrList(popRefRead, curRowNum, iDefault.doorCons, "Door Construction", defaultDoor, listDoor)
Call prrNumeric(popRefRead, curRowNum, iDefault.doorWidth, "Door Width", 3, 1, 30, "ft", "m")
Call prrNumeric(popRefRead, curRowNum, iDefault.doorHeight, "Door Height", 7, 1, 50, "ft", "m")
' style or internal gains
Call prrGroupTitle(popRefRead, curRowNum, 1, "Internal Gain Types")
styleGroupRow = curRowNum
For i = 1 To numIStyle
  Call prrGroupTitle(popRefRead, curRowNum, 2, "Internal Gains: ", iStyle(i).nm)
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).peopDensUse, "People (Operation)", iStyle(i).peopDensUse, 0, 10000, "sqft/person", "m2/person")
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).peopDensNonUse, "People (Off-Hours)", iStyle(i).peopDensNonUse, 0, 10000, "sqft/person", "m2/person")
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).liteDensUse, "Lighting (Operation)", iStyle(i).liteDensUse, 0, 10, "W/sqft", "W/m2")
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).liteDensNonUse, "Lighting (Off-Hours)", iStyle(i).liteDensNonUse, 0, 10, "W/sqft", "W/m2")
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).eqpDensUse, "Equipment (Operation)", iStyle(i).eqpDensUse, 0, 10, "W/sqft", "W/m2")
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).eqpDensNonUse, "Equipment (Off-Hours)", iStyle(i).eqpDensNonUse, 0, 10, "W/sqft", "W/m2")
  Call prrList(popRefRead, curRowNum, iStyle(i).weekdayTimeRange, "Weekday Operation", iStyle(i).weekdayTimeRange, listTimeRange)
  Call prrList(popRefRead, curRowNum, iStyle(i).saturdayTimeRange, "Saturday Operation", iStyle(i).saturdayTimeRange, listTimeRange)
  Call prrList(popRefRead, curRowNum, iStyle(i).sundayTimeRange, "Sunday-Holiday Operation", iStyle(i).sundayTimeRange, listTimeRange)
  Call prrNumeric(popRefRead, curRowNum, iStyle(i).furnDens, "Furniture Density", iStyle(i).furnDens, 5, 100, "lbs/sqft", "kg/m2")
Next i
' floorplan
Call prrGroupTitle(popRefRead, curRowNum, 1, "Floorplans")
floorplanGroupRow = curRowNum
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    Call prrGroupTitle(popRefRead, curRowNum, 2, "Floorplan: ", iFloorPlan(i).nm)
    Call prrNumeric(popRefRead, curRowNum, iFloorPlan(i).flr2flr, "Floor to Floor Height", useNumericDefault, 1, 40, "ft", "m")
    Call prrNumeric(popRefRead, curRowNum, iFloorPlan(i).numFlr, "Number of Floors", 1, 1, 40, "", "")
    Call prrShowOnlyNumeric(popRefRead, curRowNum, iFloorPlan(i).flrArea, "Floor Area", "sqft", "m2")
'    zone
    Call prrGroupTitle(popRefRead, curRowNum, 3, "Zones")
    For j = 1 To numPZone
      Call prrGroupTitle(popRefRead, curRowNum, 4, "Zone: ", pZone(j).nm(i))
      Call prrShowOnlyNumeric(popRefRead, curRowNum, pZone(j).area, "Floor Area", "sqft", "m2")
      Call prrList(popRefRead, curRowNum, pZone(j).style(i), "Internal Gains", useDefault, listStyle)
    Next j
'    extwall
    Call prrGroupTitle(popRefRead, curRowNum, 3, "Exterior Walls")
    For j = 1 To numPExtWall
      Call prrGroupTitle(popRefRead, curRowNum, 4, "Exterior Wall: ", pExtWall(j).nm(i))
      Call prrList(popRefRead, curRowNum, pExtWall(j).cons(i), "Construction", useDefault, listWallConstruction)
      Call prrList(popRefRead, curRowNum, pExtWall(j).insul(i), "Insulation", useDefault, listInsulation)
      Call prrShowOnlyNumeric(popRefRead, curRowNum, pExtWall(j).length, "Length", "ft", "m")
      Call prrShowOnlyNumeric(popRefRead, curRowNum, pExtWall(j).area(i), "Area", "sqft", "m2")
      Call prrShowOnlyNumeric(popRefRead, curRowNum, pExtWall(j).perGlaz(i), "Glazing", "%", "%")
'        window
      For k = 1 To windowsPerWall
        Call prrGroupTitle(popRefRead, curRowNum, 5, "Window: ", iWindow(i, j, k).nm)
        Call prrList(popRefRead, curRowNum, iWindow(i, j, k).cons, "Type", useDefault, listWindow)
        Call prrNumeric(popRefRead, curRowNum, iWindow(i, j, k).width, "Width", useNumericDefault, 1, 12, "ft", "m")
        Call prrNumeric(popRefRead, curRowNum, iWindow(i, j, k).height, "Height", useNumericDefault, 1, 50, "ft", "m")
        Call prrNumeric(popRefRead, curRowNum, iWindow(i, j, k).count, "Count", 0, 0, 99, "", "")
'not yet implemented Call prrNumeric(popRefRead, curRowNum, iWindow(i, j, k).ovrhng, "Overhang", useNumericDefault, 0, 20, "ft", "m")
'not yet implemented Call prrNumeric(popRefRead, curRowNum, iWindow(i, j, k).setbck, "Setback", useNumericDefault, 0, 10, "ft", "m")
      Next k
'        door
      For k = 1 To doorsPerWall
        Call prrGroupTitle(popRefRead, curRowNum, 5, "Door: ", iDoor(i, j, k).nm)
        Call prrList(popRefRead, curRowNum, iDoor(i, j, k).cons, "Construction", useDefault, listDoor)
        Call prrNumeric(popRefRead, curRowNum, iDoor(i, j, k).width, "Width", useNumericDefault, 1, 30, "ft", "m")
        Call prrNumeric(popRefRead, curRowNum, iDoor(i, j, k).height, "Height", useNumericDefault, 1, 50, "ft", "m")
        Call prrNumeric(popRefRead, curRowNum, iDoor(i, j, k).count, "Count", 0, 0, 99, "", "")
      Next k
    Next j  'extwall loop
'    intwall
    Call prrGroupTitle(popRefRead, curRowNum, 3, "Interior Walls")
    For j = 1 To numPIntWall
      Call prrGroupTitle(popRefRead, curRowNum, 5, "Interior Wall: ", pIntWall(j).nm(i))
      Call prrShowOnlyNumeric(popRefRead, curRowNum, pIntWall(j).length, "Length", "ft", "m")
      Call prrShowOnlyNumeric(popRefRead, curRowNum, pIntWall(j).area(i), "Area", "sqft", "m2")
    Next j
  End If
Next i 'floorplan loop
' corner
Call prrGroupTitle(popRefRead, curRowNum, 1, "Corners")
cornerGroupRow = curRowNum
For i = 1 To numPCorner
  Call prrGroupTitle(popRefRead, curRowNum, 2, "Corner: ", pCorner(i).name)
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pCorner(i).x, "X", "ft", "m")
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pCorner(i).y, "Y", "ft", "m")
  Call prrShowOnlyString(popRefRead, curRowNum, pCorner(i).xexpression, "X Expression")
  Call prrShowOnlyString(popRefRead, curRowNum, pCorner(i).yexpression, "Y Expression")
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pCorner(i).xTrans, "X Rotated", "ft", "m")
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pCorner(i).yTrans, "Y Rotated", "ft", "m")
Next i
' roof corner
Call prrGroupTitle(popRefRead, curRowNum, 1, "Roof Corners")
roofCornerGroupRow = curRowNum
For i = 1 To numPRoofCorner
  Call prrGroupTitle(popRefRead, curRowNum, 2, "Roof Corner: ", pRoofCorner(i).name)
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pRoofCorner(i).x, "X", "ft", "m")
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pRoofCorner(i).y, "Y", "ft", "m")
  Call prrShowOnlyString(popRefRead, curRowNum, pRoofCorner(i).xexpression, "X Expression")
  Call prrShowOnlyString(popRefRead, curRowNum, pRoofCorner(i).yexpression, "Y Expression")
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pRoofCorner(i).xTrans, "X Rotated", "ft", "m")
  Call prrShowOnlyNumeric(popRefRead, curRowNum, pRoofCorner(i).yTrans, "Y Rotated", "ft", "m")
Next i
If popRefRead = prrPopulate Or popRefRead = prrRepopulate Then
  ' collapse the entire heirarchy
  grdMain.IsCollapsed(buildingGroupRow) = flexOutlineCollapsed
  'grdMain.IsCollapsed(defaultGroupRow) = flexOutlineCollapsed
  grdMain.IsCollapsed(styleGroupRow) = flexOutlineCollapsed
  grdMain.IsCollapsed(floorplanGroupRow) = flexOutlineCollapsed
  grdMain.IsCollapsed(cornerGroupRow) = flexOutlineCollapsed
  grdMain.IsCollapsed(roofCornerGroupRow) = flexOutlineCollapsed
End If
End Sub

'------------------------------------------------------------------------
' Populate, Read or Refresh the Group Title on the grid
'------------------------------------------------------------------------
Sub prrGroupTitle(prrFlag As Integer, rowNum As Integer, indentLevel As Integer, TitleString As String, Optional SubTitleString As String)
Call incrementRow(rowNum)
If prrFlag = prrPopulate Or prrFlag = prrRepopulate Then
  grdMain.TextMatrix(rowNum, 0) = TitleString
  grdMain.TextMatrix(rowNum, 1) = ""
  grdMain.TextMatrix(rowNum, 2) = SubTitleString
  grdMain.IsSubtotal(rowNum) = True
  grdMain.Cell(flexcpFontBold, rowNum, 0, rowNum, 2) = True
  grdMain.RowOutlineLevel(rowNum) = indentLevel
  rowData(rowNum).kindOfRow = kindGroupTitle
End If
End Sub

'------------------------------------------------------------------------
' Populate, Read or Refresh a numeric input on the grid
'------------------------------------------------------------------------
Sub prrNumeric(prrFlag As Integer, rowNum As Integer, currentVal As Single, paramString As String, defaultVal As Single, minVal As Single, maxVal As Single, IPunits As String, SIunits As String)
Dim builtString As String
Call incrementRow(rowNum)
Select Case prrFlag
  Case prrPopulate
    grdMain.TextMatrix(rowNum, 0) = paramString
    If defaultVal = useNumericDefault Then
      grdMain.TextMatrix(rowNum, 2) = indicateDefault
    Else
      grdMain.TextMatrix(rowNum, 2) = Str(defaultVal)
    End If
    If newPlanInfo.isIPunits Then
      grdMain.TextMatrix(rowNum, 1) = IPunits
    Else
      grdMain.TextMatrix(rowNum, 1) = SIunits
      'force a unit conversion if SI is chosen on all length measurements
      If IPunits = "ft" Then
        If defaultVal <> useNumericDefault Then defaultVal = defaultVal * 0.3
        minVal = minVal * 0.3
        maxVal = maxVal * 0.3
      End If
    End If
    currentVal = defaultVal
    ' set up the row data
    rowData(rowNum).kindOfRow = kindNumeric
    If defaultVal = useNumericDefault Then
      builtString = "|" & indicateDefault & vbTab & "Default"
    Else
      builtString = "|" & Trim(Format(defaultVal)) & vbTab & "Default"
    End If
    builtString = builtString & "|" & Format(minVal) & vbTab & "Minimum"
    builtString = builtString & "|" & Format(maxVal) & vbTab & "Maximum"
    rowData(rowNum).defaultMinMaxString = builtString
    rowData(rowNum).defaultVal = defaultVal
    rowData(rowNum).minVal = minVal
    rowData(rowNum).maxVal = maxVal
  Case prrRepopulate
    grdMain.TextMatrix(rowNum, 0) = paramString
    If currentVal = useNumericDefault Then
      grdMain.TextMatrix(rowNum, 2) = indicateDefault
    Else
      grdMain.TextMatrix(rowNum, 2) = Str(defaultVal)
    End If
    If newPlanInfo.isIPunits Then
      grdMain.TextMatrix(rowNum, 1) = IPunits
    Else
      grdMain.TextMatrix(rowNum, 1) = SIunits
    End If
    ' set up the row data
    rowData(rowNum).kindOfRow = kindNumeric
    If defaultVal = useNumericDefault Then
      builtString = "|" & indicateDefault & vbTab & "Default"
    Else
      builtString = "|" & Trim(Format(defaultVal)) & vbTab & "Default"
    End If
    builtString = builtString & "|" & Format(minVal) & vbTab & "Minimum"
    builtString = builtString & "|" & Format(maxVal) & vbTab & "Maximum"
    rowData(rowNum).defaultMinMaxString = builtString
    rowData(rowNum).defaultVal = defaultVal
    rowData(rowNum).minVal = minVal
    rowData(rowNum).maxVal = maxVal
  Case prrRefresh
    If currentVal = useNumericDefault Then
      grdMain.TextMatrix(rowNum, 2) = indicateDefault
    Else
      grdMain.TextMatrix(rowNum, 2) = Str(currentVal)
    End If
  Case prrRead
    If grdMain.TextMatrix(rowNum, 2) = indicateDefault Then
      currentVal = useNumericDefault
    Else
      currentVal = Val(grdMain.TextMatrix(rowNum, 2))
    End If
End Select
End Sub

'------------------------------------------------------------------------
' Populate, Read or Refresh a list of choices input on the grid
'------------------------------------------------------------------------
Sub prrList(prrFlag As Integer, rowNum As Integer, ByRef currentVal As Integer, paramString As String, defaultVal As Integer, ListKind As Integer)
Dim curSelItem As String
Dim Found As Integer
Dim i As Integer
Call incrementRow(rowNum)
Select Case prrFlag
  Case prrPopulate
    grdMain.TextMatrix(rowNum, 0) = paramString
    grdMain.TextMatrix(rowNum, 1) = ""
    grdMain.TextMatrix(rowNum, 2) = listOfChoices(defaultVal)
    currentVal = defaultVal
    rowData(rowNum).kindOfRow = kindList
    If defaultVal = useDefault Then
      rowData(rowNum).listOfOptionsString = "Use Default|" & kindOfList(ListKind).builtString
    Else
      rowData(rowNum).listOfOptionsString = kindOfList(ListKind).builtString
    End If
  Case prrRepopulate
    grdMain.TextMatrix(rowNum, 0) = paramString
    grdMain.TextMatrix(rowNum, 1) = ""
    grdMain.TextMatrix(rowNum, 2) = listOfChoices(currentVal)
    rowData(rowNum).kindOfRow = kindList
    If defaultVal = useDefault Then
      rowData(rowNum).listOfOptionsString = "Use Default|" & kindOfList(ListKind).builtString
    Else
      rowData(rowNum).listOfOptionsString = kindOfList(ListKind).builtString
    End If
  Case prrRefresh
    grdMain.TextMatrix(rowNum, 2) = listOfChoices(currentVal)
  Case prrRead
    'search through the list of options for the string of text in the current cell
    Found = 0
    curSelItem = grdMain.TextMatrix(rowNum, 2)
    If curSelItem = indicateDefault Then
      currentVal = useDefault
    Else
      For i = kindOfList(ListKind).firstChoice To kindOfList(ListKind).lastChoice
        If curSelItem = listOfChoices(i) Then
          Found = i
          Exit For
        End If
      Next i
      'if a match is found then assign the index to the current value
      If Found > 0 Then
        currentVal = Found
      Else
        currentVal = defaultVal
        MsgBox "Could not find selected list item in list: " & curSelItem, vbCritical, "Error"
      End If
    End If
End Select
End Sub

'------------------------------------------------------------------------
' Populate or Refresh a numeric show only parameter on the grid
'------------------------------------------------------------------------
Sub prrShowOnlyNumeric(prrFlag As Integer, rowNum As Integer, currentVal As Single, paramString As String, IPunits As String, SIunits As String)
Call incrementRow(rowNum)
Select Case prrFlag
  Case prrPopulate, prrRepopulate
    grdMain.TextMatrix(rowNum, 0) = paramString
    grdMain.TextMatrix(rowNum, 2) = Str(currentVal)
    grdMain.Cell(flexcpForeColor, rowNum, 2, rowNum, 2) = vbBlue
    If newPlanInfo.isIPunits Then
      grdMain.TextMatrix(rowNum, 1) = IPunits
    Else
      grdMain.TextMatrix(rowNum, 1) = SIunits
    End If
    rowData(rowNum).kindOfRow = kindShowOnlyNumeric
  Case prrRefresh
    grdMain.TextMatrix(rowNum, 2) = Str(currentVal)
  Case prrRead
    'do nothing - show only parameter
End Select
End Sub

'------------------------------------------------------------------------
' Populate or Refresh a string show only parameter on the grid
'------------------------------------------------------------------------
Sub prrShowOnlyString(prrFlag As Integer, rowNum As Integer, currentVal As String, paramString As String)
Call incrementRow(rowNum)
Select Case prrFlag
  Case prrPopulate, prrRepopulate
    grdMain.TextMatrix(rowNum, 0) = paramString
    grdMain.TextMatrix(rowNum, 1) = ""
    grdMain.TextMatrix(rowNum, 2) = currentVal
    grdMain.Cell(flexcpForeColor, rowNum, 2, rowNum, 2) = vbBlue
    rowData(rowNum).kindOfRow = kindShowOnlyString
  Case prrRefresh
    grdMain.TextMatrix(rowNum, 2) = currentVal
  Case prrRead
    'do nothing - show only parameter
End Select
End Sub

'------------------------------------------------------------------------
' Increment the counter when populating the grid
'------------------------------------------------------------------------
Sub incrementRow(rowIndx As Integer)
If rowIndx < grdMain.Rows Then
  rowIndx = rowIndx + 1
Else
  MsgBox "Increment Row call causes index to be greater than the number of rows", vbCritical, "Error Message"
End If
End Sub

'------------------------------------------------------------------------
' The list of possible choices for input parameters is initialized
'------------------------------------------------------------------------
Sub initializeChoiceList()
listOfChoices(0) = indicateDefault
'----- wall constructions
listWallConstruction = 1
defaultWallConstruction = 1
kindOfList(listWallConstruction).firstChoice = 1
kindOfList(listWallConstruction).lastChoice = 47
listOfChoices(1) = "Face Brick + Ins + 4in LW Concrete Block + Gyp"
constLayer(1).nm = "FaceBrkIns4LWConcBlkGyp"
constLayer(1).matCount = 4
constLayer(1).matInd(1) = 2
constLayer(1).matInd(2) = constLayerInsul
constLayer(1).matInd(3) = 19
constLayer(1).matInd(4) = 34
listOfChoices(2) = "Face Brick + Ins + 8in LW Concrete Block + Gyp"
constLayer(2).nm = "FaceBrkIns8LWConcBlkGyp"
constLayer(2).matCount = 4
constLayer(2).matInd(1) = 2
constLayer(2).matInd(2) = constLayerInsul
constLayer(2).matInd(3) = 24
constLayer(2).matInd(4) = 34
listOfChoices(3) = "Face Brick + Ins + 2in HW Concrete Block + Gyp"
constLayer(3).nm = "FaceBrkIns2HWConcBlkGyp"
constLayer(3).matCount = 4
constLayer(3).matInd(1) = 2
constLayer(3).matInd(2) = constLayerInsul
constLayer(3).matInd(3) = 29
constLayer(3).matInd(4) = 34
listOfChoices(4) = "Face Brick + Ins + 4in HW Concrete Block + Gyp"
constLayer(4).nm = "FaceBrkIns4HWConcBlkGyp"
constLayer(4).matCount = 4
constLayer(4).matInd(1) = 2
constLayer(4).matInd(2) = constLayerInsul
constLayer(4).matInd(3) = 22
constLayer(4).matInd(4) = 34
listOfChoices(5) = "Face Brick + Ins + 8in HW Concrete Block + Gyp"
constLayer(5).nm = "FaceBrkIns8HWConcBlkGyp"
constLayer(5).matCount = 4
constLayer(5).matInd(1) = 2
constLayer(5).matInd(2) = constLayerInsul
constLayer(5).matInd(3) = 25
constLayer(5).matInd(4) = 34
listOfChoices(6) = "Face Brick + Ins + 4in Common Brick + Gyp"
constLayer(6).nm = "FaceBrkIns4CommonBrkGyp"
constLayer(6).matCount = 4
constLayer(6).matInd(1) = 2
constLayer(6).matInd(2) = constLayerInsul
constLayer(6).matInd(3) = 21
constLayer(6).matInd(4) = 34
listOfChoices(7) = "Face Brick + Ins + 8in Common Brick + Gyp"
constLayer(7).nm = "FaceBrkIns8CommonBrkGyp"
constLayer(7).matCount = 4
constLayer(7).matInd(1) = 2
constLayer(7).matInd(2) = constLayerInsul
constLayer(7).matInd(3) = 26
constLayer(7).matInd(4) = 34
listOfChoices(8) = "Face Brick + Ins + 8in Clay Tile + Gyp"
constLayer(8).nm = "FaceBrkIns8ClayTileGyp"
constLayer(8).matCount = 4
constLayer(8).matInd(1) = 2
constLayer(8).matInd(2) = constLayerInsul
constLayer(8).matInd(3) = 23
constLayer(8).matInd(4) = 34
listOfChoices(9) = "Face Brick + Ins + 8in HW Concrete + Gyp"
constLayer(9).nm = ""
constLayer(9).matCount = 4
constLayer(9).matInd(1) = 2
constLayer(9).matInd(2) = constLayerInsul
constLayer(9).matInd(3) = 27
constLayer(9).matInd(4) = 34
listOfChoices(10) = "Face Brick + Ins + 4in LW Concrete + Gyp"
constLayer(10).nm = "FaceBrkIns4LWConcGyp"
constLayer(10).matCount = 4
constLayer(10).matInd(1) = 2
constLayer(10).matInd(2) = constLayerInsul
constLayer(10).matInd(3) = 31
constLayer(10).matInd(4) = 34
listOfChoices(11) = "Face Brick + Ins + 12in HW Concrete + Gyp"
constLayer(11).nm = "FaceBrkIns12HWConcGyp"
constLayer(11).matCount = 4
constLayer(11).matInd(1) = 2
constLayer(11).matInd(2) = constLayerInsul
constLayer(11).matInd(3) = 28
constLayer(11).matInd(4) = 34
listOfChoices(12) = "Face Brick + 4in LW Concrete Block + Ins + Gyp"
constLayer(12).nm = "FaceBrk4LWConcBlkInsGyp"
constLayer(12).matCount = 4
constLayer(12).matInd(1) = 2
constLayer(12).matInd(2) = 19
constLayer(12).matInd(3) = constLayerInsul
constLayer(12).matInd(4) = 34
listOfChoices(13) = "Face Brick + 8in LW Concrete Block + Ins + Gyp"
constLayer(13).nm = "FaceBrk8LWConcBlkInsGyp"
constLayer(13).matCount = 4
constLayer(13).matInd(1) = 2
constLayer(13).matInd(2) = 24
constLayer(13).matInd(3) = constLayerInsul
constLayer(13).matInd(4) = 34
listOfChoices(14) = "Face Brick + 2in HW Concrete Block + Ins + Gyp"
constLayer(14).nm = "FaceBrk2HWConcBlkInsGyp"
constLayer(14).matCount = 4
constLayer(14).matInd(1) = 2
constLayer(14).matInd(2) = 29
constLayer(14).matInd(3) = constLayerInsul
constLayer(14).matInd(4) = 34
listOfChoices(15) = "Face Brick + 4in HW Concrete Block + Ins + Gyp"
constLayer(15).nm = "FaceBrk4HWConcBlockInsGyp"
constLayer(15).matCount = 4
constLayer(15).matInd(1) = 2
constLayer(15).matInd(2) = 20
constLayer(15).matInd(3) = constLayerInsul
constLayer(15).matInd(4) = 34
listOfChoices(16) = "Face Brick + 8in HW Concrete Block + Ins + Gyp"
constLayer(16).nm = "FaceBrk8HWConcBlkInsGyp"
constLayer(16).matCount = 4
constLayer(16).matInd(1) = 2
constLayer(16).matInd(2) = 25
constLayer(16).matInd(3) = constLayerInsul
constLayer(16).matInd(4) = 34
listOfChoices(17) = "Face Brick + 4in Common Brick + Ins + Gyp"
constLayer(17).nm = "FaceBrk4CommonBrkInsGyp"
constLayer(17).matCount = 4
constLayer(17).matInd(1) = 2
constLayer(17).matInd(2) = 21
constLayer(17).matInd(3) = constLayerInsul
constLayer(17).matInd(4) = 34
listOfChoices(18) = "Face Brick + 8in Common Brick + Ins + Gyp"
constLayer(18).nm = "FaceBrk8CommonBrkInsGyp"
constLayer(18).matCount = 4
constLayer(18).matInd(1) = 2
constLayer(18).matInd(2) = 26
constLayer(18).matInd(3) = constLayerInsul
constLayer(18).matInd(4) = 34
listOfChoices(19) = "Face Brick + 8in Clay Tile + Ins + Gyp"
constLayer(19).nm = "FaceBrk8ClayTileInsGyp"
constLayer(19).matCount = 4
constLayer(19).matInd(1) = 2
constLayer(19).matInd(2) = 23
constLayer(19).matInd(3) = constLayerInsul
constLayer(19).matInd(4) = 34
listOfChoices(20) = "Face Brick + 8in HW Concrete + Ins + Gyp"
constLayer(20).nm = "FaceBrk8HWConcInsGyp"
constLayer(20).matCount = 4
constLayer(20).matInd(1) = 2
constLayer(20).matInd(2) = 27
constLayer(20).matInd(3) = constLayerInsul
constLayer(20).matInd(4) = 34
listOfChoices(21) = "Face Brick + 4in LW Concrete + Ins + Gyp"
constLayer(21).nm = "FaceBrk4LWConcInsGyp"
constLayer(21).matCount = 4
constLayer(21).matInd(1) = 2
constLayer(21).matInd(2) = 31
constLayer(21).matInd(3) = constLayerInsul
constLayer(21).matInd(4) = 34
listOfChoices(22) = "Face Brick + 12in HW Concrete + Ins + Gyp"
constLayer(22).nm = "FaceBrk12HWConcInsGyp"
constLayer(22).matCount = 4
constLayer(22).matInd(1) = 2
constLayer(22).matInd(2) = 28
constLayer(22).matInd(3) = constLayerInsul
constLayer(22).matInd(4) = 34
listOfChoices(23) = "4in LW Concrete Block + Ins + Gyp"
constLayer(23).nm = "LW4ConcBlkInsGyp"
constLayer(23).matCount = 3
constLayer(23).matInd(1) = 19
constLayer(23).matInd(2) = constLayerInsul
constLayer(23).matInd(3) = 34
listOfChoices(24) = "8in LW Concrete Block + Ins + Gyp"
constLayer(24).nm = "LW8ConcBlkInsGyp"
constLayer(24).matCount = 3
constLayer(24).matInd(1) = 24
constLayer(24).matInd(2) = constLayerInsul
constLayer(24).matInd(3) = 34
listOfChoices(25) = "2in HW Concrete Block + Ins + Gyp"
constLayer(25).nm = "HW2ConcBlkInsGyp"
constLayer(25).matCount = 3
constLayer(25).matInd(1) = 29
constLayer(25).matInd(2) = constLayerInsul
constLayer(25).matInd(3) = 34
listOfChoices(26) = "4in HW Concrete Block + Ins + Gyp"
constLayer(26).nm = "HW4ConcBlkInsGyp"
constLayer(26).matCount = 3
constLayer(26).matInd(1) = 20
constLayer(26).matInd(2) = constLayerInsul
constLayer(26).matInd(3) = 34
listOfChoices(27) = "8in HW Concrete Block + Ins + Gyp"
constLayer(27).nm = "HW8ConcBlkInsGyp"
constLayer(27).matCount = 3
constLayer(27).matInd(1) = 25
constLayer(27).matInd(2) = constLayerInsul
constLayer(27).matInd(3) = 34
listOfChoices(28) = "4in Common Brick + Ins + Gyp"
constLayer(28).nm = "CommonBrk4InsGyp"
constLayer(28).matCount = 3
constLayer(28).matInd(1) = 21
constLayer(28).matInd(2) = constLayerInsul
constLayer(28).matInd(3) = 34
listOfChoices(29) = "8in Common Brick + Ins + Gyp"
constLayer(29).nm = "CommonBrk8InsGyp"
constLayer(29).matCount = 3
constLayer(29).matInd(1) = 26
constLayer(29).matInd(2) = constLayerInsul
constLayer(29).matInd(3) = 34
listOfChoices(30) = "8in Clay Tile + Ins + Gyp"
constLayer(30).nm = "ClayTile8InsGyp"
constLayer(30).matCount = 3
constLayer(30).matInd(1) = 23
constLayer(30).matInd(2) = constLayerInsul
constLayer(30).matInd(3) = 34
listOfChoices(31) = "8in HW Concrete + Ins + Gyp"
constLayer(31).nm = "HW8ConcInsGyp"
constLayer(31).matCount = 3
constLayer(31).matInd(1) = 27
constLayer(31).matInd(2) = constLayerInsul
constLayer(31).matInd(3) = 34
listOfChoices(32) = "4in LW Concrete + Ins + Gyp"
constLayer(32).nm = "LW4ConcInsGyp"
constLayer(32).matCount = 3
constLayer(32).matInd(1) = 31
constLayer(32).matInd(2) = constLayerInsul
constLayer(32).matInd(3) = 34
listOfChoices(33) = "12in HW Concrete + Ins + Gyp"
constLayer(33).nm = "HW12ConcInsGyp"
constLayer(33).matCount = 3
constLayer(33).matInd(1) = 28
constLayer(33).matInd(2) = constLayerInsul
constLayer(33).matInd(3) = 34
listOfChoices(34) = "Stucco + 4in LW Concrete Block + Ins + Gyp"
constLayer(34).nm = "Stucco4LWConcBlkInsGyp"
constLayer(34).matCount = 4
constLayer(34).matInd(1) = 1
constLayer(34).matInd(2) = 19
constLayer(34).matInd(3) = constLayerInsul
constLayer(34).matInd(4) = 34
listOfChoices(35) = "Stucco + 8in LW Concrete Block + Ins + Gyp"
constLayer(35).nm = "Stucco8LWConcBlkInsGyp"
constLayer(35).matCount = 4
constLayer(35).matInd(1) = 1
constLayer(35).matInd(2) = 24
constLayer(35).matInd(3) = constLayerInsul
constLayer(35).matInd(4) = 34
listOfChoices(36) = "Stucco + 2in HW Concrete Block + Ins + Gyp"
constLayer(36).nm = "Stucco2HWConcBlkInsGyp"
constLayer(36).matCount = 4
constLayer(36).matInd(1) = 1
constLayer(36).matInd(2) = 29
constLayer(36).matInd(3) = constLayerInsul
constLayer(36).matInd(4) = 34
listOfChoices(37) = "Stucco + 4in HW Concrete Block + Ins + Gyp"
constLayer(37).nm = "Stucco4HWConcBlkInsGyp"
constLayer(37).matCount = 4
constLayer(37).matInd(1) = 1
constLayer(37).matInd(2) = 20
constLayer(37).matInd(3) = constLayerInsul
constLayer(37).matInd(4) = 34
listOfChoices(38) = "Stucco + 8in HW Concrete Block + Ins + Gyp"
constLayer(38).nm = "Stucco8HWConcBlkInsGyp"
constLayer(38).matCount = 4
constLayer(38).matInd(1) = 1
constLayer(38).matInd(2) = 25
constLayer(38).matInd(3) = constLayerInsul
constLayer(38).matInd(4) = 34
listOfChoices(39) = "Stucco + 4in Common Brick + Ins + Gyp"
constLayer(39).nm = "Stucco4CommonBrkInsGyp"
constLayer(39).matCount = 4
constLayer(39).matInd(1) = 1
constLayer(39).matInd(2) = 21
constLayer(39).matInd(3) = constLayerInsul
constLayer(39).matInd(4) = 34
listOfChoices(40) = "Stucco + 8in Common Brick + Ins + Gyp"
constLayer(40).nm = "Stucco8CommonBrkInsGyp"
constLayer(40).matCount = 4
constLayer(40).matInd(1) = 1
constLayer(40).matInd(2) = 26
constLayer(40).matInd(3) = constLayerInsul
constLayer(40).matInd(4) = 34
listOfChoices(41) = "Stucco + 8in Clay Tile + Ins + Gyp"
constLayer(41).nm = "Stucco8ClayTileInsGyp"
constLayer(41).matCount = 4
constLayer(41).matInd(1) = 1
constLayer(41).matInd(2) = 23
constLayer(41).matInd(3) = constLayerInsul
constLayer(41).matInd(4) = 34
listOfChoices(42) = "Stucco + 8in HW Concrete + Ins + Gyp"
constLayer(42).nm = "Stucco8HWConcInsGyp"
constLayer(42).matCount = 4
constLayer(42).matInd(1) = 1
constLayer(42).matInd(2) = 25
constLayer(42).matInd(3) = constLayerInsul
constLayer(42).matInd(4) = 34
listOfChoices(43) = "Stucco + 4in LW Concrete + Ins + Gyp"
constLayer(43).nm = "Stucco4LWConcInsGyp"
constLayer(43).matCount = 4
constLayer(43).matInd(1) = 1
constLayer(43).matInd(2) = 19
constLayer(43).matInd(3) = constLayerInsul
constLayer(43).matInd(4) = 34
listOfChoices(44) = "Stucco + 12in HW Concrete + Ins + Gyp"
constLayer(44).nm = "Stucco12HWConcInsGyp"
constLayer(44).matCount = 4
constLayer(44).matInd(1) = 1
constLayer(44).matInd(2) = 28
constLayer(44).matInd(3) = constLayerInsul
constLayer(44).matInd(4) = 34
listOfChoices(45) = "Steel + Ins + Steel"
constLayer(45).nm = "SteelInsSteel"
constLayer(45).matCount = 3
constLayer(45).matInd(1) = 3
constLayer(45).matInd(2) = constLayerInsul
constLayer(45).matInd(3) = 3
listOfChoices(46) = "Steel + Ins + 4in HW Concrete + Gyp"
constLayer(46).nm = "SteelIns4HWConcGyp"
constLayer(46).matCount = 4
constLayer(46).matInd(1) = 3
constLayer(46).matInd(2) = constLayerInsul
constLayer(46).matInd(3) = 22
constLayer(46).matInd(4) = 34
listOfChoices(47) = "Steel + Ins + 8in HW Concrete + Gyp"
constLayer(47).nm = "SteelIns8HWConcGyp"
constLayer(47).matCount = 4
constLayer(47).matInd(1) = 3
constLayer(47).matInd(2) = constLayerInsul
constLayer(47).matInd(3) = 27
constLayer(47).matInd(4) = 34
'----- roof constructions
listRoofConstruction = 2
defaultRoofConstruction = 50
kindOfList(listRoofConstruction).firstChoice = 48
kindOfList(listRoofConstruction).lastChoice = 63
listOfChoices(48) = "Slag + Felt + Ins + Wood + Gap + AcoTile"
constLayer(48).nm = "SlagFeltInsWoodGapAcoTile"
constLayer(48).matCount = 6
constLayer(48).matInd(1) = 35
constLayer(48).matInd(2) = 36
constLayer(48).matInd(3) = constLayerInsul
constLayer(48).matInd(4) = 12
constLayer(48).matInd(5) = constLayerAirGap  'CEILING AIR SPACE  MATERIAL:Air,HF-E4,0.1762000;
constLayer(48).matInd(6) = 38
listOfChoices(49) = "Slag + Felt + Ins + 4in LW Concrete + Gap + AcoTile"
constLayer(49).nm = "SlagFeltIns4LWConcGapAcoTile"
constLayer(49).matCount = 6
constLayer(49).matInd(1) = 35
constLayer(49).matInd(2) = 36
constLayer(49).matInd(3) = constLayerInsul
constLayer(49).matInd(4) = 31
constLayer(49).matInd(5) = constLayerAirGap
constLayer(49).matInd(6) = 38
listOfChoices(50) = "Slag + Felt + Ins + 6in LW Concrete + Gap + AcoTile"
constLayer(50).nm = "SlagFeltIns6LWConcGapAcoTile"
constLayer(50).matCount = 6
constLayer(50).matInd(1) = 35
constLayer(50).matInd(2) = 36
constLayer(50).matInd(3) = constLayerInsul
constLayer(50).matInd(4) = 32
constLayer(50).matInd(5) = constLayerAirGap
constLayer(50).matInd(6) = 38
listOfChoices(51) = "Slag + Felt + Ins + 8in LW Concrete + Gap + AcoTile"
constLayer(51).nm = "SlagFeltIns8LWConcGapAcoTile"
constLayer(51).matCount = 6
constLayer(51).matInd(1) = 35
constLayer(51).matInd(2) = 36
constLayer(51).matInd(3) = constLayerInsul
constLayer(51).matInd(4) = 33
constLayer(51).matInd(5) = constLayerAirGap
constLayer(51).matInd(6) = 38
listOfChoices(52) = "Slag + Felt + Ins + 2in HW Concrete + Gap + AcoTile"
constLayer(52).nm = "SlagFeltIns2HWConcGapAcoTile"
constLayer(52).matCount = 6
constLayer(52).matInd(1) = 35
constLayer(52).matInd(2) = 36
constLayer(52).matInd(3) = constLayerInsul
constLayer(52).matInd(4) = 29
constLayer(52).matInd(5) = constLayerAirGap
constLayer(52).matInd(6) = 38
listOfChoices(53) = "Slag + Felt + Ins + 4in HW Concrete + Gap + AcoTile"
constLayer(53).nm = "SlagFeltIns4HWConcGapAcoTile"
constLayer(53).matCount = 6
constLayer(53).matInd(1) = 35
constLayer(53).matInd(2) = 36
constLayer(53).matInd(3) = constLayerInsul
constLayer(53).matInd(4) = 22
constLayer(53).matInd(5) = constLayerAirGap
constLayer(53).matInd(6) = 38
listOfChoices(54) = "Slag + Felt + Ins + 6in HW Concrete + Gap + AcoTile"
constLayer(54).nm = "SlagFeltIns6HWConcGapAcoTile"
constLayer(54).matCount = 6
constLayer(54).matInd(1) = 35
constLayer(54).matInd(2) = 36
constLayer(54).matInd(3) = constLayerInsul
constLayer(54).matInd(4) = 30
constLayer(54).matInd(5) = constLayerAirGap
constLayer(54).matInd(6) = 38
listOfChoices(55) = "Slag + Felt + Ins + Steel Siding + Gap + AcoTile"
constLayer(55).nm = "SlagFeltInsSteelSidingGapAcoTile"
constLayer(55).matCount = 6
constLayer(55).matInd(1) = 35
constLayer(55).matInd(2) = 36
constLayer(55).matInd(3) = constLayerInsul
constLayer(55).matInd(4) = 3
constLayer(55).matInd(5) = constLayerAirGap
constLayer(55).matInd(6) = 38
listOfChoices(56) = "Slag + Felt + Ins + Wood"
constLayer(56).nm = "SlagFeltInsWood"
constLayer(56).matCount = 4
constLayer(56).matInd(1) = 35
constLayer(56).matInd(2) = 36
constLayer(56).matInd(3) = constLayerInsul
constLayer(56).matInd(4) = 13
listOfChoices(57) = "Slag + Felt + Ins + 4in LW Concrete"
constLayer(57).nm = "SlagFeltIns4LWConc"
constLayer(57).matCount = 4
constLayer(57).matInd(1) = 35
constLayer(57).matInd(2) = 36
constLayer(57).matInd(3) = constLayerInsul
constLayer(57).matInd(4) = 31
listOfChoices(58) = "Slag + Felt + Ins + 6in LW Concrete"
constLayer(58).nm = "SlagFeltIns6LWConc"
constLayer(58).matCount = 4
constLayer(58).matInd(1) = 35
constLayer(58).matInd(2) = 36
constLayer(58).matInd(3) = constLayerInsul
constLayer(58).matInd(4) = 32
listOfChoices(59) = "Slag + Felt + Ins + 8in LW Concrete"
constLayer(59).nm = "SlagFeltIns8LWConc"
constLayer(59).matCount = 4
constLayer(59).matInd(1) = 35
constLayer(59).matInd(2) = 36
constLayer(59).matInd(3) = constLayerInsul
constLayer(59).matInd(4) = 33
listOfChoices(60) = "Slag + Felt + Ins + 2in HW Concrete"
constLayer(60).nm = "SlagFeltIns2HWConc"
constLayer(60).matCount = 4
constLayer(60).matInd(1) = 35
constLayer(60).matInd(2) = 36
constLayer(60).matInd(3) = constLayerInsul
constLayer(60).matInd(4) = 29
listOfChoices(61) = "Slag + Felt + Ins + 4in HW Concrete"
constLayer(61).nm = "SlagFeltIns4HW Conc"
constLayer(61).matCount = 4
constLayer(61).matInd(1) = 35
constLayer(61).matInd(2) = 36
constLayer(61).matInd(3) = constLayerInsul
constLayer(61).matInd(4) = 22
listOfChoices(62) = "Slag + Felt + Ins + 6in HW Concrete"
constLayer(62).nm = "SlagFeltIns6HWConc"
constLayer(62).matCount = 4
constLayer(62).matInd(1) = 35
constLayer(62).matInd(2) = 36
constLayer(62).matInd(3) = constLayerInsul
constLayer(62).matInd(4) = 30
listOfChoices(63) = "Slag + Felt + Ins + Steel Siding"
constLayer(63).nm = "SlagFeltInsSteelSiding"
constLayer(63).matCount = 4
constLayer(63).matInd(1) = 35
constLayer(63).matInd(2) = 36
constLayer(63).matInd(3) = constLayerInsul
constLayer(63).matInd(4) = 3
'----- insulation
listInsulation = 3
defaultInsulation = 68
kindOfList(listInsulation).firstChoice = 64
kindOfList(listInsulation).lastChoice = 87
listOfChoices(64) = "R-3"
insulation(1).rValue = 3
listOfChoices(65) = "R-5"
insulation(2).rValue = 5
listOfChoices(66) = "R-7"
insulation(3).rValue = 6
listOfChoices(67) = "R-9"
insulation(4).rValue = 7
listOfChoices(68) = "R-11"
insulation(5).rValue = 11
listOfChoices(69) = "R-13"
insulation(6).rValue = 13
listOfChoices(70) = "R-17"
insulation(7).rValue = 17
listOfChoices(71) = "R-19"
insulation(8).rValue = 19
listOfChoices(72) = "R-21"
insulation(9).rValue = 21
listOfChoices(73) = "R-23"
insulation(10).rValue = 23
listOfChoices(74) = "R-25"
insulation(11).rValue = 25
listOfChoices(75) = "R-27"
insulation(12).rValue = 27
listOfChoices(76) = "R-29"
insulation(13).rValue = 29
listOfChoices(77) = "R-31"
insulation(14).rValue = 31
listOfChoices(78) = "R-33"
insulation(15).rValue = 33
listOfChoices(79) = "R-35"
insulation(16).rValue = 35
listOfChoices(80) = "R-37"
insulation(17).rValue = 37
listOfChoices(81) = "R-39"
insulation(18).rValue = 39
listOfChoices(82) = "R-41"
insulation(19).rValue = 41
listOfChoices(83) = "R-43"
insulation(20).rValue = 43
listOfChoices(84) = "R-47"
insulation(21).rValue = 47
listOfChoices(85) = "R-51"
insulation(22).rValue = 51
listOfChoices(86) = "R-55"
insulation(23).rValue = 55
listOfChoices(87) = "R-59"
insulation(24).rValue = 59
'----- window
listWindow = 4
defaultWindow = 140
kindOfList(listWindow).firstChoice = 100
kindOfList(listWindow).lastChoice = 306
listOfChoices(100) = "1000 - Sgl Clr 3mm "
listOfChoices(101) = "1001 - Sgl Clr 6mm "
listOfChoices(102) = "1002 - Sgl Clr Low Iron 3mm "
listOfChoices(103) = "1003 - Sgl Clr Low Iron 5mm "
listOfChoices(104) = "1200 - Sgl Bronze 3mm "
listOfChoices(105) = "1201 - Sgl Bronze 6mm "
listOfChoices(106) = "1202 - Sgl Green 3mm "
listOfChoices(107) = "1203 - Sgl Green 6mm "
listOfChoices(108) = "1204 - Sgl Grey 3mm "
listOfChoices(109) = "1205 - Sgl Grey 6mm "
listOfChoices(110) = "1206 - Sgl Blue 6mm "
listOfChoices(111) = "1400 - Sgl Ref-A-L Clr 6mm "
listOfChoices(112) = "1401 - Sgl Ref-A-M Clr 6mm "
listOfChoices(113) = "1402 - Sgl Ref-A-H Clr 6mm "
listOfChoices(114) = "1403 - Sgl Ref-A-L Tint 6mm "
listOfChoices(115) = "1404 - Sgl Ref-A-M Tint 6mm "
listOfChoices(116) = "1405 - Sgl Ref-A-H Tint 6mm "
listOfChoices(117) = "1406 - Sgl Ref-B-L Clr 6mm "
listOfChoices(118) = "1407 - Sgl Ref-B-H Clr 6mm "
listOfChoices(119) = "1408 - Sgl Ref-B-L Tint 6mm "
listOfChoices(120) = "1409 - Sgl Ref-B-M Tint 6mm "
listOfChoices(121) = "1410 - Sgl Ref-B-H Tint 6mm "
listOfChoices(122) = "1411 - Sgl Ref-C-L Clr 6mm "
listOfChoices(123) = "1412 - Sgl Ref-C-M Clr 6mm "
listOfChoices(124) = "1413 - Sgl Ref-C-H Clr 6mm "
listOfChoices(125) = "1414 - Sgl Ref-C-L Tint 6mm "
listOfChoices(126) = "1415 - Sgl Ref-C-M Tint 6mm "
listOfChoices(127) = "1416 - Sgl Ref-C-H Tint 6mm "
listOfChoices(128) = "1417 - Sgl Ref-D Clr 6mm "
listOfChoices(129) = "1418 - Sgl Ref-D Tint 6mm "
listOfChoices(130) = "1600 - Sgl LoE (e2=.4) Clr 3mm "
listOfChoices(131) = "1601 - Sgl LoE (e2=.2) Clr 3mm "
listOfChoices(132) = "1602 - Sgl LoE (e2=.2) Clr 6mm "
listOfChoices(133) = "1800 - Sgl Elec Abs Bleached 6mm "
listOfChoices(134) = "1801 - Sgl Elec Abs Colored 6mm "
listOfChoices(135) = "1802 - Sgl Elec Ref Bleached 6mm "
listOfChoices(136) = "1803 - Sgl Elec Ref Colored 6mm "
listOfChoices(137) = "2000 - Dbl Clr 3mm/6mm Air "
listOfChoices(138) = "2001 - Dbl Clr 3mm/13mm Air "
listOfChoices(139) = "2002 - Dbl Clr 3mm/13mm Arg "
listOfChoices(140) = "2003 - Dbl Clr 6mm/6mm Air "
listOfChoices(141) = "2004 - Dbl Clr 6mm/13mm Air "
listOfChoices(142) = "2005 - Dbl Clr 6mm/13mm Arg "
listOfChoices(143) = "2006 - Dbl Clr Low Iron 3mm/6mm Air "
listOfChoices(144) = "2007 - Dbl Clr Low Iron 3mm/13mm Air "
listOfChoices(145) = "2008 - Dbl Clr Low Iron 3mm/13mm Arg "
listOfChoices(146) = "2009 - Dbl Clr Low Iron 5mm/6mm Air "
listOfChoices(147) = "2010 - Dbl Clr Low Iron 5mm/13mm Air "
listOfChoices(148) = "2011 - Dbl Clr Low Iron 5mm/13mm Arg "
listOfChoices(149) = "2200 - Dbl Bronze 3mm/6mm Air "
listOfChoices(150) = "2201 - Dbl Bronze 3mm/13mm Air "
listOfChoices(151) = "2202 - Dbl Bronze 3mm/13mm Arg "
listOfChoices(152) = "2203 - Dbl Bronze 6mm/6mm Air "
listOfChoices(153) = "2204 - Dbl Bronze 6mm/13mm Air "
listOfChoices(154) = "2205 - Dbl Bronze 6mm/13mm Arg "
listOfChoices(155) = "2206 - Dbl Green 3m/6mm Air "
listOfChoices(156) = "2207 - Dbl Green 3mm/13mm Air "
listOfChoices(157) = "2208 - Dbl Green 3mm/13mm Arg "
listOfChoices(158) = "2209 - Dbl Green 6mm/6mm Air "
listOfChoices(159) = "2210 - Dbl Green 6mm/13mm Air "
listOfChoices(160) = "2211 - Dbl Green 6mm/13mm Arg "
listOfChoices(161) = "2212 - Dbl Grey 3mm/6mm Air "
listOfChoices(162) = "2213 - Dbl Grey 3mm/13mm Air "
listOfChoices(163) = "2214 - Dbl Grey 3mm/13mm Arg "
listOfChoices(164) = "2215 - Dbl Grey 6mm/6mm Air "
listOfChoices(165) = "2216 - Dbl Grey 6mm/13mm Air "
listOfChoices(166) = "2217 - Dbl Grey 6mm/13mm Arg "
listOfChoices(167) = "2218 - Dbl Blue 6mm/6mm Air "
listOfChoices(168) = "2219 - Dbl Blue 6mm/13mm Air "
listOfChoices(169) = "2220 - Dbl Blue 6mm/13mm Arg "
listOfChoices(170) = "2400 - Dbl Ref-A-L Clr 6mm/6mm Air "
listOfChoices(171) = "2401 - Dbl Ref-A-L Clr 6mm/13mm Air "
listOfChoices(172) = "2402 - Dbl Ref-A-L Clr 6mm/13mm Arg "
listOfChoices(173) = "2403 - Dbl Ref-A-M Clr 6mm/6mm Air "
listOfChoices(174) = "2404 - Dbl Ref-A-M Clr 6mm/13mm Air "
listOfChoices(175) = "2405 - Dbl Ref-A-M Clr 6mm/13mm Arg "
listOfChoices(176) = "2406 - Dbl Ref-A-H Clr 6mm/6mm Air "
listOfChoices(177) = "2407 - Dbl Ref-A-H 6mm/13mm Air "
listOfChoices(178) = "2408 - Dbl Ref-A-H Clr 6mm/13mm Arg "
listOfChoices(179) = "2410 - Dbl Ref-A-L Tint 6mm/6mm Air "
listOfChoices(180) = "2411 - Dbl Ref-A-L Tint 6mm/13mm Air "
listOfChoices(181) = "2412 - Dbl Ref-A-L Tint 6mm/13mm Arg "
listOfChoices(182) = "2413 - Dbl Ref-A-M Tint 6mm/6mm Air "
listOfChoices(183) = "2414 - Dbl Ref-A-M Tint 6mm/13mm Air "
listOfChoices(184) = "2415 - Dbl Ref-A-M Tint 6mm/13mm Arg "
listOfChoices(185) = "2416 - Dbl Ref-A-H Tint 6mm/6mm Air "
listOfChoices(186) = "2417 - Dbl Ref-A-H Tint 6mm/13mm Air "
listOfChoices(187) = "2418 - Dbl Ref-A-H Tint 6mm/13mm Arg "
listOfChoices(188) = "2420 - Dbl Ref-B-L Clr 6mm/6mm Air "
listOfChoices(189) = "2421 - Dbl Ref-B-L Clr 6mm/13mm Air "
listOfChoices(190) = "2422 - Dbl Ref-B-L Clr 6mm/13mm Arg "
listOfChoices(191) = "2426 - Dbl Ref-B-H Clr 6mm/6mm Air "
listOfChoices(192) = "2427 - Dbl Ref-B-H Clr 6mm/13mm Air "
listOfChoices(193) = "2428 - Dbl Ref-B-H Clr 6mm/13mm Arg "
listOfChoices(194) = "2430 - Dbl Ref-B-L Tint 6mm/6mm Air "
listOfChoices(195) = "2431 - Dbl Ref-B-L Tint 6mm/13mm Air "
listOfChoices(196) = "2432 - Dbl Ref-B-L Tint 6mm/13mm Arg "
listOfChoices(197) = "2433 - Dbl Ref-B-M Tint 6mm/6mm Air "
listOfChoices(198) = "2434 - Dbl Ref-B-M Tint 6mm/13mm Air "
listOfChoices(199) = "2435 - Dbl Ref-B-M Tint 6mm/13mm Arg "
listOfChoices(200) = "2436 - Dbl Ref-B-H Tint 6mm/6mm Air "
listOfChoices(201) = "2437 - Dbl Ref-B-H Tint 6mm/13mm Air "
listOfChoices(202) = "2438 - Dbl Ref B-H Tint 6mm/13mm Arg "
listOfChoices(203) = "2440 - Dbl Ref-C-L Clr 6mm/6mm Air "
listOfChoices(204) = "2441 - Dbl Ref-C-L Clr 6mm/13mm Air "
listOfChoices(205) = "2442 - Dbl Ref-C-L Clr 6mm/13mm Arg "
listOfChoices(206) = "2443 - Dbl Ref-C-M Clr 6mm/6mm Air "
listOfChoices(207) = "2444 - Dbl Ref-C-M Clr 6mm/13mm Air "
listOfChoices(208) = "2445 - Dbl Ref-C-M Clr 6mm/13mm Arg "
listOfChoices(209) = "2446 - Dbl Ref-C-H Clr 6mm/6mm Air "
listOfChoices(210) = "2447 - Dbl Ref-C-H Clr 6mm/13mm Air "
listOfChoices(211) = "2448 - Dbl Ref-C-H Clr 6mm/13mm Arg "
listOfChoices(212) = "2450 - Dbl Ref-C-L Tint 6mm/6mm Air "
listOfChoices(213) = "2451 - Dbl Ref-C-L Tint 6mm/13mm Air "
listOfChoices(214) = "2452 - Dbl Ref-C-L Tint 6mm/13mm Arg "
listOfChoices(215) = "2453 - Dbl Ref-C-M Tint 6mm/6mm Air "
listOfChoices(216) = "2454 - Dbl Ref-C-M Tint 6mm/13mm Air "
listOfChoices(217) = "2455 - Dbl Ref-C-M Tint 6mm/13mm Arg "
listOfChoices(218) = "2456 - Dbl Ref-C-H Tint 6mm/6mm Air "
listOfChoices(219) = "2457 - Dbl Ref-C-H Tint 6mm/13mm Air "
listOfChoices(220) = "2458 - Dbl Ref-C-H Tint 6mm/13mm Arg "
listOfChoices(221) = "2460 - Dbl Ref-D Clr 6mm/6mm Air "
listOfChoices(222) = "2461 - Dbl Ref-D Clr 6mm/13mm Air "
listOfChoices(223) = "2462 - Dbl Ref-D Clr 6mm/13mm Arg "
listOfChoices(224) = "2470 - Dbl Ref-D Tint 6mm/6mm Air "
listOfChoices(225) = "2471 - Dbl Ref-D Tint 6mm/13mm Air "
listOfChoices(226) = "2472 - Dbl Ref-D Tint 6mm/13mm Arg "
listOfChoices(227) = "2600 - Dbl LoE (e2=.4) Clr 3mm/6mm Air "
listOfChoices(228) = "2601 - Dbl LoE (e2=.4) Clr 3mm/13mm Air "
listOfChoices(229) = "2602 - Dbl LoE (e2=.4) Clr 3mm/13mm Arg "
listOfChoices(230) = "2610 - Dbl LoE (e2=.2) Clr 3mm/6mm Air "
listOfChoices(231) = "2611 - Dbl LoE (e2=.2) Clr 3mm/13mm Air "
listOfChoices(232) = "2612 - Dbl LoE (e2=.2) Clr 3mm/13mm Arg "
listOfChoices(233) = "2613 - Dbl LoE (e2=.2) Clr 6mm/6mm Air "
listOfChoices(234) = "2614 - Dbl LoE (e2=.2) Clr 6mm/13mm Air "
listOfChoices(235) = "2615 - Dbl LoE (e2=.2) Clr 6mm/13mm Arg "
listOfChoices(236) = "2630 - Dbl LoE (e2=.1) Clr 3mm/6mm Air "
listOfChoices(237) = "2631 - Dbl LoE (e2=.1) Clr 3mm/13mm Air "
listOfChoices(238) = "2632 - Dbl LoE (e2=.1) Clr 3mm/13mm Arg "
listOfChoices(239) = "2633 - Dbl LoE (e2=.1) Clr 6mm/6mm Air "
listOfChoices(240) = "2634 - Dbl LoE (e2=.1) Clr 6mm/13mm Air "
listOfChoices(241) = "2635 - Dbl LoE (e2=.1) Clr 6mm/13mm Arg "
listOfChoices(242) = "2636 - Dbl LoE (e2=.1) Tint 6mm/6mm Air "
listOfChoices(243) = "2637 - Dbl LoE (e2=.1) Tint 6mm/13mm Air "
listOfChoices(244) = "2638 - Dbl LoE (e2=.1) Tint 6mm/13mm Arg "
listOfChoices(245) = "2640 - Dbl LoE (e3=.1) Clr 3mm/6mm Air "
listOfChoices(246) = "2641 - Dbl LoE (e3=.1) Clr 3mm/13mm Air "
listOfChoices(247) = "2642 - Dbl LoE (e3=.1) Clr 3mm/13mm Arg "
listOfChoices(248) = "2660 - Dbl LoE Spec Sel Clr 3mm/6mm/6mm Air "
listOfChoices(249) = "2661 - Dbl LoE Spec Sel Clr 3mm/13mm/6mm Air "
listOfChoices(250) = "2662 - Dbl LoE Spec Sel Clr 3mm/13mm/6mm Arg "
listOfChoices(251) = "2663 - Dbl LoE Spec Sel Clr 6mm/6mm Air "
listOfChoices(252) = "2664 - Dbl LoE Spec Sel Clr 6mm/13mm Air "
listOfChoices(253) = "2665 - Dbl LoE Spec Sel Clr 6mm/13mm Arg "
listOfChoices(254) = "2666 - Dbl LoE Spec Sel Tint 6mm/6mm Air "
listOfChoices(255) = "2667 - Dbl LoE Spec Sel Tint 6mm/13mm Air "
listOfChoices(256) = "2668 - Dbl LoE Spec Sel Tint 6mm/13mm Arg "
listOfChoices(257) = "2800 - Dbl Elec Abs Bleached 6mm/6mm Air "
listOfChoices(258) = "2801 - Dbl Elec Abs Colored 6mm/6mm Air "
listOfChoices(259) = "2802 - Dbl Elec Abs Bleached 6mm/13mm Air "
listOfChoices(260) = "2803 - Dbl Elec Abs Colored 6mm/13mm Air "
listOfChoices(261) = "2804 - Dbl Elec Abs Bleached 6mm/13mm Arg "
listOfChoices(262) = "2805 - Dbl Elec Abs Colored 6mm/13mm Arg "
listOfChoices(263) = "2820 - Dbl Elec Ref Bleached 6mm/6mm Air "
listOfChoices(264) = "2821 - Dbl Elec Ref Colored 6mm/6mm Air "
listOfChoices(265) = "2822 - Dbl Elec Ref Bleached 6mm/13mm Air "
listOfChoices(266) = "2823 - Dbl Elec Ref Colored 6mm/13mm Air "
listOfChoices(267) = "2824 - Dbl Elec Ref Bleached 6mm/13mm Arg "
listOfChoices(268) = "2825 - Dbl Elec Ref Colored 6mm/13mm Arg "
listOfChoices(269) = "2840 - Dbl LoE Elec Abs Bleached 6mm/6mm Air "
listOfChoices(270) = "2841 - Dbl LoE Elec Abs Colored 6mm/6mm Air "
listOfChoices(271) = "2842 - Dbl LoE Elec Abs Bleached 6mm/13mm Air "
listOfChoices(272) = "2843 - Dbl LoE Elec Abs Colored 6mm/13mm Air "
listOfChoices(273) = "2844 - Dbl LoE Elec Abs Bleached 6mm/13mm Arg "
listOfChoices(274) = "2845 - Dbl LoE Elec Abs Colored 6mm/13mm Arg "
listOfChoices(275) = "2860 - Dbl LoE Elec Ref Bleached 6mm/6mm Air "
listOfChoices(276) = "2861 - Dbl LoE Elec Ref Colored 6mm/6mm Air "
listOfChoices(277) = "2862 - Dbl LoE Elec Ref Bleached 6mm/13mm Air "
listOfChoices(278) = "2863 - Dbl LoE Elec Ref Colored 6mm/13mm Air "
listOfChoices(279) = "2864 - Dbl LoE Elec Ref Bleached 6mm/13mm Arg "
listOfChoices(280) = "2865 - Dbl LoE Elec Ref Colored 6mm/13m Arg "
listOfChoices(281) = "3001 - Trp Clr 3mm/6mm Air "
listOfChoices(282) = "3002 - Trp Clr 3mm/13mm Air "
listOfChoices(283) = "3003 - Trp Clr 3mm/13mm Arg "
listOfChoices(284) = "3601 - Trp LoE (e5=.1) Clr 3mm/6mm Air "
listOfChoices(285) = "3602 - Trp LoE (e5=.1) Clr 3mm/13mm Air "
listOfChoices(286) = "3603 - Trp LoE (e5=.1) Clr 3mm/13mm Arg "
listOfChoices(287) = "3621 - Trp LoE (e2=e5=.1) Clr 3mm/6mm Air "
listOfChoices(288) = "3622 - Trp LoE (e2=e5=.1) Clr 3mm/13mm Air "
listOfChoices(289) = "3623 - Trp LoE (e2=e5=.1) Clr 3mm/13mm Arg "
listOfChoices(290) = "3641 - Trp LoE Film (88) Clr 3mm/6mm Air "
listOfChoices(291) = "3642 - Trp LoE Film (88) Clr 3mm/13mm Air "
listOfChoices(292) = "3651 - Trp LoE Film (77) Clr 3mm/6mm Air "
listOfChoices(293) = "3652 - Trp LoE Film (77) Clr 3mm/13mm Air "
listOfChoices(294) = "3661 - Trp LoE Film (66) Clr 6mm/6mm Air "
listOfChoices(295) = "3662 - Trp LoE Film (66) Clr 6mm/13mm Air "
listOfChoices(296) = "3663 - Trp LoE Film (66) Bronze 6mm/6mm Air "
listOfChoices(297) = "3664 - Trp LoE Film (66) Bronze 6mm/13mm Air "
listOfChoices(298) = "3671 - Trp LoE Film (55) Clr 6mm/6mm Air "
listOfChoices(299) = "3672 - Trp LoE Film (55) Clr 6mm/13m Air "
listOfChoices(300) = "3673 - Trp LoE Film (55) Bronze 6mm/6mm Air "
listOfChoices(301) = "3674 - Trp LoE Film (55) Bronze 6mm/13mm Air "
listOfChoices(302) = "3681 - Trp LoE Film (44) Bronze 6mm/6mm Air "
listOfChoices(303) = "3682 - Trp LoE Film (44) Bronze 6mm/13mm Air "
listOfChoices(304) = "3691 - Trp LoE Film (33) Bronze 6mm/6mm Air "
listOfChoices(305) = "3692 - Trp LoE Film (33) Bronze 6mm/13mm Air "
listOfChoices(306) = "4651 - Quadruple LoE Films (88) 3mm/8mm Krypton"
'----- door
listDoor = 5
defaultDoor = 324
kindOfList(listDoor).firstChoice = 320
kindOfList(listDoor).lastChoice = 327
listOfChoices(320) = "Aluminum"
listOfChoices(321) = "Aluminum Roll - up"
listOfChoices(322) = "Glass"
listOfChoices(323) = "Hollow wood"
listOfChoices(324) = "Metal Insulated"
listOfChoices(325) = "Sliding Partition"
listOfChoices(326) = "Solid Wood"
listOfChoices(327) = "Wood Roll-up"
'----- schedule
listSchedule = 6
defaultSchedule = 339
kindOfList(listSchedule).firstChoice = 330
kindOfList(listSchedule).lastChoice = 343
listOfChoices(330) = "typical 8 to 6"
listOfChoices(331) = "24/7"
listOfChoices(332) = "Assembly"
listOfChoices(333) = "Health"
listOfChoices(334) = "Hotel"
listOfChoices(335) = "Manufacturing"
listOfChoices(336) = "Multifamily"
listOfChoices(337) = "Multifamily OneZone"
listOfChoices(338) = "Multifamily TwoZoneBedroom"
listOfChoices(339) = "Office"
listOfChoices(340) = "Restaurant"
listOfChoices(341) = "Retail"
listOfChoices(342) = "School"
listOfChoices(343) = "Warehouse"
'----- floor constructions
listFloorConstruction = 7
defaultFloorConstruction = 354
kindOfList(listFloorConstruction).firstChoice = 350
kindOfList(listFloorConstruction).lastChoice = 355
listOfChoices(350) = "4in LW Concrete"
listOfChoices(351) = "6in LW Concrete"
listOfChoices(352) = "8in LW Concrete"
listOfChoices(353) = "4in HW Concrete"
listOfChoices(354) = "6in HW Concrete"
listOfChoices(355) = "8in HW Concrete"
'----- style
listStyle = 8
defaultStyle = 360
kindOfList(listStyle).firstChoice = 360
kindOfList(listStyle).lastChoice = 369
listOfChoices(360) = "office"
listOfChoices(361) = "retail"
listOfChoices(362) = "residential"
listOfChoices(363) = "conference"
listOfChoices(364) = "medical"
listOfChoices(365) = "storage"
listOfChoices(366) = "dining"
listOfChoices(367) = "auditorium"
listOfChoices(368) = "kitchen"
listOfChoices(369) = "gym"
'----- Interior wall constructions
listIntWallCons = 9
defaultIntWallCons = 380
kindOfList(listIntWallCons).firstChoice = 380
kindOfList(listIntWallCons).lastChoice = 382
listOfChoices(380) = "gyp + gap + concrete + gap + gyp"
listOfChoices(381) = "concrete"
listOfChoices(382) = "gyp + wood frame + gyp"
'----- Time Ranges
listTimeRange = 10
defaultTimeRange = 412
kindOfList(listTimeRange).firstChoice = 400
kindOfList(listTimeRange).lastChoice = 440
listOfChoices(400) = "Hours 10am to 4pm"
listOfChoices(401) = "Hours 10am to 5pm"
listOfChoices(402) = "Hours 10am to 6pm"
listOfChoices(403) = "Hours 10am to 7pm"
listOfChoices(404) = "Hours 10am to 8pm"
listOfChoices(405) = "Hours 9am to 4pm"
listOfChoices(406) = "Hours 9am to 5pm"
listOfChoices(407) = "Hours 9am to 6pm"
listOfChoices(408) = "Hours 9am to 7pm"
listOfChoices(409) = "Hours 9am to 8pm"
listOfChoices(410) = "Hours 8am to 4pm"
listOfChoices(411) = "Hours 8am to 5pm"
listOfChoices(412) = "Hours 8am to 6pm"
listOfChoices(413) = "Hours 8am to 7pm"
listOfChoices(414) = "Hours 8am to 8pm"
listOfChoices(415) = "Hours 7am to 4pm"
listOfChoices(416) = "Hours 7am to 5pm"
listOfChoices(417) = "Hours 7am to 6pm"
listOfChoices(418) = "Hours 7am to 7pm"
listOfChoices(419) = "Hours 7am to 8pm"
listOfChoices(420) = "Hours 6am to 4pm"
listOfChoices(421) = "Hours 6am to 5pm"
listOfChoices(422) = "Hours 6am to 6pm"
listOfChoices(423) = "Hours 6am to 7pm"
listOfChoices(424) = "Hours 6am to 8pm"
listOfChoices(425) = "Hours 5am to 4pm"
listOfChoices(426) = "Hours 5am to 5pm"
listOfChoices(427) = "Hours 5am to 6pm"
listOfChoices(428) = "Hours 5am to 7pm"
listOfChoices(429) = "Hours 5am to 8pm"
listOfChoices(430) = "Hours 11am to 3pm"
listOfChoices(431) = "Hours 11am to 4pm"
listOfChoices(432) = "Hours noon to 2pm"
listOfChoices(433) = "Hours noon to 3pm"
listOfChoices(434) = "Hours 4am to 8pm"
listOfChoices(435) = "Hours 4am to 9pm"
listOfChoices(436) = "Hours 3am to 9pm"
listOfChoices(437) = "Hours 3am to 10pm"
listOfChoices(438) = "Hours 2am to 10pm"
listOfChoices(439) = "Hours 2am to 11pm"
listOfChoices(440) = "All Hours"
'----- Durations
listDuration = 11
defaultDuration = 480
kindOfList(listDuration).firstChoice = 480
kindOfList(listDuration).lastChoice = 481
listOfChoices(480) = "Annual"
durationAnnual = 480
listOfChoices(481) = "Design Days"
durationDesign = 481
'------ EnergyPlus Version Names
listEPversion = 12
defaultEPversion = 510
kindOfList(listEPversion).firstChoice = 500
kindOfList(listEPversion).lastChoice = 510
listOfChoices(500) = "1.2.1"
epVersion121 = 500
listOfChoices(501) = "1.2.2"
epVersion122 = 501
listOfChoices(502) = "1.2.3"
epVersion123 = 5020
listOfChoices(503) = "1.3"
epVersion130 = 503
listOfChoices(504) = "1.4.0"
epVersion140 = 504
listOfChoices(505) = "2.0.0"
epVersion200 = 505
listOfChoices(506) = "2.1.0"
epVersion210 = 506
listOfChoices(507) = "2.2.0"
epVersion220 = 507
listOfChoices(508) = "3.0.0"
epVersion300 = 508
listOfChoices(509) = "3.1.0"
epVersion310 = 509
listOfChoices(510) = "4.0.0"
epVersion400 = 510
End Sub

'------------------------------------------------------------------------
' Build up choice list strings
'------------------------------------------------------------------------
Sub buildChoiceListString()
Dim st As String
Dim i As Integer, j As Integer
For i = 1 To maxNumKindOfList
  st = ""
  For j = kindOfList(i).firstChoice To kindOfList(i).lastChoice
    st = st & listOfChoices(j)
    If j < kindOfList(i).lastChoice Then
      st = st & "|"
    End If
  Next j
  kindOfList(i).builtString = st
Next i
End Sub

'------------------------------------------------------------------------
' Initialize the values in the style array - must be consistant with
' the listOfChoices array
'------------------------------------------------------------------------
Sub initializeStyle()
Dim ft2m As Single
Dim sqft2sqm As Single
Dim btuhPerSqft2WPerSqm As Single
Dim lbPerSqft2kgPerSqm As Single
Dim i As Integer
For i = 1 To numIStyle
  iStyle(i).isUsed = False
  iStyle(i).weekdayTimeRange = 412 '8 to 6
  iStyle(i).saturdayTimeRange = 412 '8 to 6
  iStyle(i).sundayTimeRange = 400 '10 to 4
Next i
iStyle(1).nm = "office"
iStyle(1).peopDensUse = 250
iStyle(1).peopDensNonUse = 5000
iStyle(1).liteDensUse = 1.75
iStyle(1).liteDensNonUse = 0.25
iStyle(1).eqpDensUse = 0.75
iStyle(1).eqpDensNonUse = 0.25
iStyle(1).furnDens = 10

iStyle(2).nm = "retail"
iStyle(2).peopDensUse = 300
iStyle(2).peopDensNonUse = 5000
iStyle(2).liteDensUse = 2.5
iStyle(2).liteDensNonUse = 0.25
iStyle(2).eqpDensUse = 0.25
iStyle(2).eqpDensNonUse = 0.05
iStyle(2).furnDens = 10

iStyle(3).nm = "residential"
iStyle(3).peopDensUse = 100
iStyle(3).peopDensNonUse = 300
iStyle(3).liteDensUse = 1.5
iStyle(3).liteDensNonUse = 0.25
iStyle(3).eqpDensUse = 0.6
iStyle(3).eqpDensNonUse = 0.1
iStyle(3).furnDens = 10

iStyle(4).nm = "conference"
iStyle(4).peopDensUse = 50
iStyle(4).peopDensNonUse = 5000
iStyle(4).liteDensUse = 1.75
iStyle(4).liteDensNonUse = 0.25
iStyle(4).eqpDensUse = 0.1
iStyle(4).eqpDensNonUse = 0.05
iStyle(4).furnDens = 10

iStyle(5).nm = "medical"
iStyle(5).peopDensUse = 300
iStyle(5).peopDensNonUse = 1000
iStyle(5).liteDensUse = 2#
iStyle(5).liteDensNonUse = 0.5
iStyle(5).eqpDensUse = 1.2
iStyle(5).eqpDensNonUse = 0.25
iStyle(5).furnDens = 10

iStyle(6).nm = "storage"
iStyle(6).peopDensUse = 2000
iStyle(6).peopDensNonUse = 20000
iStyle(6).liteDensUse = 1.3
iStyle(6).liteDensNonUse = 0.1
iStyle(6).eqpDensUse = 0.1
iStyle(6).eqpDensNonUse = 0.05
iStyle(6).furnDens = 10

iStyle(7).nm = "dining"
iStyle(7).peopDensUse = 50
iStyle(7).peopDensNonUse = 5000
iStyle(7).liteDensUse = 2.2
iStyle(7).liteDensNonUse = 0.25
iStyle(7).eqpDensUse = 0.1
iStyle(7).eqpDensNonUse = 0.05
iStyle(7).furnDens = 10

iStyle(8).nm = "auditorium"
iStyle(8).peopDensUse = 50
iStyle(8).peopDensNonUse = 5000
iStyle(8).liteDensUse = 1.3
iStyle(8).liteDensNonUse = 0.25
iStyle(8).eqpDensUse = 0.5
iStyle(8).eqpDensNonUse = 0.05
iStyle(8).furnDens = 10

iStyle(9).nm = "kitchen"
iStyle(9).peopDensUse = 200
iStyle(9).peopDensNonUse = 1000
iStyle(9).liteDensUse = 2.2
iStyle(9).liteDensNonUse = 0.25
iStyle(9).eqpDensUse = 3#
iStyle(9).eqpDensNonUse = 0.5
iStyle(9).furnDens = 10

iStyle(10).nm = "gym"
iStyle(10).peopDensUse = 140
iStyle(10).peopDensNonUse = 2000
iStyle(10).liteDensUse = 1.9
iStyle(10).liteDensNonUse = 0.25
iStyle(10).eqpDensUse = 0.1
iStyle(10).eqpDensNonUse = 0.01
iStyle(10).furnDens = 10

'convert styles to metric if necessary
If Not newPlanInfo.isIPunits Then
  ft2m = 0.3  ' approx 1 / 3.281
  sqft2sqm = 0.1  'approx 1 / (3.281 * 3.281)
  btuhPerSqft2WPerSqm = 3 'approx 1 / 0.316957210776545
  lbPerSqft2kgPerSqm = 5   'approx (3.281 * 3.281) / 2.2 from 2.2 kg/lb and 3.281 ft/m
  For i = 1 To numIStyle
    iStyle(i).peopDensUse = iStyle(i).peopDensUse * sqft2sqm
    iStyle(i).peopDensNonUse = iStyle(i).peopDensNonUse * sqft2sqm
    iStyle(i).liteDensUse = iStyle(i).liteDensUse * btuhPerSqft2WPerSqm
    iStyle(i).liteDensNonUse = iStyle(i).liteDensNonUse * btuhPerSqft2WPerSqm
    iStyle(i).eqpDensUse = iStyle(i).eqpDensUse * btuhPerSqft2WPerSqm
    iStyle(i).eqpDensNonUse = iStyle(i).eqpDensNonUse * btuhPerSqft2WPerSqm
    iStyle(i).furnDens = iStyle(i).furnDens * lbPerSqft2kgPerSqm
  Next i
End If
End Sub

'------------------------------------------------------------------------
' Initialize the values in the materials array
'------------------------------------------------------------------------
Sub initializeMaterials()
Dim iLay As Integer
MATERIAL(1).nm = "HF-A1"
MATERIAL(1).desc = "STUCCO 1IN"
MATERIAL(1).rough = 1
MATERIAL(1).thick = 0.0253
MATERIAL(1).conduct = 0.6918
MATERIAL(1).dens = 1858
MATERIAL(1).spheat = 837
MATERIAL(1).emit = 0.9
MATERIAL(1).solAbs = 0.92
MATERIAL(1).visAbs = 0.92
MATERIAL(2).nm = "HF-A2"
MATERIAL(2).desc = "FACE BRICK 4IN"
MATERIAL(2).rough = 1
MATERIAL(2).thick = 0.1016
MATERIAL(2).conduct = 1.332
MATERIAL(2).dens = 2082
MATERIAL(2).spheat = 920
MATERIAL(2).emit = 0.9
MATERIAL(2).solAbs = 0.93
MATERIAL(2).visAbs = 0.93
MATERIAL(3).nm = "HF-A3"
MATERIAL(3).desc = "STEEL SIDING LW"
MATERIAL(3).rough = 1
MATERIAL(3).thick = 0.0015
MATERIAL(3).conduct = 44.97
MATERIAL(3).dens = 7689
MATERIAL(3).spheat = 418
MATERIAL(3).emit = 0.9
MATERIAL(3).solAbs = 0.2
MATERIAL(3).visAbs = 0.2
MATERIAL(4).nm = "HF-A6"
MATERIAL(4).desc = "FINISH"
MATERIAL(4).rough = 1
MATERIAL(4).thick = 0.0127
MATERIAL(4).conduct = 0.4151
MATERIAL(4).dens = 1249
MATERIAL(4).spheat = 1088
MATERIAL(4).emit = 0.9
MATERIAL(4).solAbs = 0.5
MATERIAL(4).visAbs = 0.5
MATERIAL(5).nm = "HF-A7"
MATERIAL(5).desc = "FACE BRICK 4IN"
MATERIAL(5).rough = 1
MATERIAL(5).thick = 0.1016
MATERIAL(5).conduct = 1.332
MATERIAL(5).dens = 2002
MATERIAL(5).spheat = 920
MATERIAL(5).emit = 0.9
MATERIAL(5).solAbs = 0.93
MATERIAL(5).visAbs = 0.93


MATERIAL(7).nm = "HF-B2"
MATERIAL(7).desc = "INSULATION 1IN"
MATERIAL(7).rough = 1
MATERIAL(7).thick = 0.0253
MATERIAL(7).conduct = 0.0432
MATERIAL(7).dens = 32
MATERIAL(7).spheat = 837
MATERIAL(7).emit = 0.9
MATERIAL(7).solAbs = 0.5
MATERIAL(7).visAbs = 0.5
MATERIAL(8).nm = "HF-B3"
MATERIAL(8).desc = "INSULATION 2IN"
MATERIAL(8).rough = 1
MATERIAL(8).thick = 0.0509
MATERIAL(8).conduct = 0.0432
MATERIAL(8).dens = 32
MATERIAL(8).spheat = 837
MATERIAL(8).emit = 0.9
MATERIAL(8).solAbs = 0.5
MATERIAL(8).visAbs = 0.5
MATERIAL(9).nm = "HF-B4"
MATERIAL(9).desc = "INSULATION 3IN"
MATERIAL(9).rough = 1
MATERIAL(9).thick = 0.0762
MATERIAL(9).conduct = 0.0432
MATERIAL(9).dens = 32
MATERIAL(9).spheat = 837
MATERIAL(9).emit = 0.9
MATERIAL(9).solAbs = 0.5
MATERIAL(9).visAbs = 0.5
MATERIAL(10).nm = "HF-B5"
MATERIAL(10).desc = "INSULATION 1IN"
MATERIAL(10).rough = 1
MATERIAL(10).thick = 0.0254
MATERIAL(10).conduct = 0.0432
MATERIAL(10).dens = 91
MATERIAL(10).spheat = 837
MATERIAL(10).emit = 0.9
MATERIAL(10).solAbs = 0.5
MATERIAL(10).visAbs = 0.5
MATERIAL(11).nm = "HF-B6"
MATERIAL(11).desc = "INSULATION 2IN"
MATERIAL(11).rough = 1
MATERIAL(11).thick = 0.0509
MATERIAL(11).conduct = 0.0432
MATERIAL(11).dens = 91
MATERIAL(11).spheat = 837
MATERIAL(11).emit = 0.9
MATERIAL(11).solAbs = 0.5
MATERIAL(11).visAbs = 0.5
MATERIAL(12).nm = "HF-B7"
MATERIAL(12).desc = "WOOD  1IN"
MATERIAL(12).rough = 1
MATERIAL(12).thick = 0.0254
MATERIAL(12).conduct = 0.1211
MATERIAL(12).dens = 593
MATERIAL(12).spheat = 2510
MATERIAL(12).emit = 0.9
MATERIAL(12).solAbs = 0.78
MATERIAL(12).visAbs = 0.78
MATERIAL(13).nm = "HF-B8"
MATERIAL(13).desc = "WOOD  2.5IN"
MATERIAL(13).rough = 1
MATERIAL(13).thick = 0.0635
MATERIAL(13).conduct = 0.1211
MATERIAL(13).dens = 593
MATERIAL(13).spheat = 2510
MATERIAL(13).emit = 0.9
MATERIAL(13).solAbs = 0.78
MATERIAL(13).visAbs = 0.78
MATERIAL(14).nm = "HF-B9"
MATERIAL(14).desc = "WOOD  4IN"
MATERIAL(14).rough = 1
MATERIAL(14).thick = 0.1016
MATERIAL(14).conduct = 0.1211
MATERIAL(14).dens = 593
MATERIAL(14).spheat = 2510
MATERIAL(14).emit = 0.9
MATERIAL(14).solAbs = 0.78
MATERIAL(14).visAbs = 0.78
MATERIAL(15).nm = "HF-B10"
MATERIAL(15).desc = "WOOD 2IN"
MATERIAL(15).rough = 1
MATERIAL(15).thick = 0.0508
MATERIAL(15).conduct = 0.1211
MATERIAL(15).dens = 593
MATERIAL(15).spheat = 2510
MATERIAL(15).emit = 0.9
MATERIAL(15).solAbs = 0.78
MATERIAL(15).visAbs = 0.78
MATERIAL(16).nm = "HF-B11"
MATERIAL(16).desc = "WOOD 3IN"
MATERIAL(16).rough = 1
MATERIAL(16).thick = 0.0762
MATERIAL(16).conduct = 0.1211
MATERIAL(16).dens = 593
MATERIAL(16).spheat = 2510
MATERIAL(16).emit = 0.9
MATERIAL(16).solAbs = 0.78
MATERIAL(16).visAbs = 0.78
MATERIAL(17).nm = "HF-B12"
MATERIAL(17).desc = "INSULATION 3IN"
MATERIAL(17).rough = 1
MATERIAL(17).thick = 0.0762
MATERIAL(17).conduct = 0.0432
MATERIAL(17).dens = 91
MATERIAL(17).spheat = 837
MATERIAL(17).emit = 0.9
MATERIAL(17).solAbs = 0.5
MATERIAL(17).visAbs = 0.5
MATERIAL(18).nm = "HF-C1"
MATERIAL(18).desc = "CLAY TILE 4IN"
MATERIAL(18).rough = 1
MATERIAL(18).thick = 0.1015
MATERIAL(18).conduct = 0.5708
MATERIAL(18).dens = 1121
MATERIAL(18).spheat = 837
MATERIAL(18).emit = 0.9
MATERIAL(18).solAbs = 0.82
MATERIAL(18).visAbs = 0.82
MATERIAL(19).nm = "HF-C2"
MATERIAL(19).desc = "CONCRETE BLOCK LW 4IN"
MATERIAL(19).rough = 1
MATERIAL(19).thick = 0.1015
MATERIAL(19).conduct = 0.3805
MATERIAL(19).dens = 609
MATERIAL(19).spheat = 837
MATERIAL(19).emit = 0.9
MATERIAL(19).solAbs = 0.65
MATERIAL(19).visAbs = 0.65
MATERIAL(20).nm = "HF-C3"
MATERIAL(20).desc = "CONCRETE BLOCK HW 4IN"
MATERIAL(20).rough = 1
MATERIAL(20).thick = 0.1016
MATERIAL(20).conduct = 0.8129
MATERIAL(20).dens = 977
MATERIAL(20).spheat = 837
MATERIAL(20).emit = 0.9
MATERIAL(20).solAbs = 0.65
MATERIAL(20).visAbs = 0.65
MATERIAL(21).nm = "HF-C4"
MATERIAL(21).desc = "COMMON BRICK 4IN"
MATERIAL(21).rough = 1
MATERIAL(21).thick = 0.1016
MATERIAL(21).conduct = 0.7264
MATERIAL(21).dens = 1922
MATERIAL(21).spheat = 837
MATERIAL(21).emit = 0.9
MATERIAL(21).solAbs = 0.76
MATERIAL(21).visAbs = 0.76
MATERIAL(22).nm = "HF-C5"
MATERIAL(22).desc = "CONCRETE HW 4IN"
MATERIAL(22).rough = 1
MATERIAL(22).thick = 0.1015
MATERIAL(22).conduct = 1.7296
MATERIAL(22).dens = 2243
MATERIAL(22).spheat = 837
MATERIAL(22).emit = 0.9
MATERIAL(22).solAbs = 0.65
MATERIAL(22).visAbs = 0.65
MATERIAL(23).nm = "HF-C6"
MATERIAL(23).desc = "CLAY TILE 8IN"
MATERIAL(23).rough = 1
MATERIAL(23).thick = 0.2033
MATERIAL(23).conduct = 0.5708
MATERIAL(23).dens = 1121
MATERIAL(23).spheat = 837
MATERIAL(23).emit = 0.9
MATERIAL(23).solAbs = 0.82
MATERIAL(23).visAbs = 0.82
MATERIAL(24).nm = "HF-C7"
MATERIAL(24).desc = "CONCRETE BLOCK LW  8IN"
MATERIAL(24).rough = 1
MATERIAL(24).thick = 0.2033
MATERIAL(24).conduct = 0.5708
MATERIAL(24).dens = 609
MATERIAL(24).spheat = 837
MATERIAL(24).emit = 0.9
MATERIAL(24).solAbs = 0.65
MATERIAL(24).visAbs = 0.65
concForInterior = 24
MATERIAL(25).nm = "HF-C8"
MATERIAL(25).desc = "CONCRETE BLOCK HW  8IN"
MATERIAL(25).rough = 1
MATERIAL(25).thick = 0.2033
MATERIAL(25).conduct = 1.038
MATERIAL(25).dens = 977
MATERIAL(25).spheat = 837
MATERIAL(25).emit = 0.9
MATERIAL(25).solAbs = 0.65
MATERIAL(25).visAbs = 0.65
MATERIAL(26).nm = "HF-C9"
MATERIAL(26).desc = "COMMON BRICK 8IN"
MATERIAL(26).rough = 1
MATERIAL(26).thick = 0.2033
MATERIAL(26).conduct = 0.7264
MATERIAL(26).dens = 1922
MATERIAL(26).spheat = 837
MATERIAL(26).emit = 0.9
MATERIAL(26).solAbs = 0.72
MATERIAL(26).visAbs = 0.72
MATERIAL(27).nm = "HF-C10"
MATERIAL(27).desc = "CONCRETE HW 8IN"
MATERIAL(27).rough = 1
MATERIAL(27).thick = 0.2033
MATERIAL(27).conduct = 1.73
MATERIAL(27).dens = 2243
MATERIAL(27).spheat = 837
MATERIAL(27).emit = 0.9
MATERIAL(27).solAbs = 0.65
MATERIAL(27).visAbs = 0.65
MATERIAL(28).nm = "HF-C11"
MATERIAL(28).desc = "CONCRETE HW 12IN"
MATERIAL(28).rough = 1
MATERIAL(28).thick = 0.3048
MATERIAL(28).conduct = 1.73
MATERIAL(28).dens = 2243
MATERIAL(28).spheat = 837
MATERIAL(28).emit = 0.9
MATERIAL(28).solAbs = 0.65
MATERIAL(28).visAbs = 0.65
MATERIAL(29).nm = "HF-C12"
MATERIAL(29).desc = "CONCRETE HW 2IN"
MATERIAL(29).rough = 1
MATERIAL(29).thick = 0.0509
MATERIAL(29).conduct = 1.73
MATERIAL(29).dens = 2243
MATERIAL(29).spheat = 837
MATERIAL(29).emit = 0.9
MATERIAL(29).solAbs = 0.65
MATERIAL(29).visAbs = 0.65
MATERIAL(30).nm = "HF-C13"
MATERIAL(30).desc = "CONCRETE HW 6IN"
MATERIAL(30).rough = 1
MATERIAL(30).thick = 0.1524
MATERIAL(30).conduct = 1.73
MATERIAL(30).dens = 2243
MATERIAL(30).spheat = 837
MATERIAL(30).emit = 0.9
MATERIAL(30).solAbs = 0.65
MATERIAL(30).visAbs = 0.65
MATERIAL(31).nm = "HF-C14"
MATERIAL(31).desc = "CONCRETE LW 4IN"
MATERIAL(31).rough = 1
MATERIAL(31).thick = 0.1016
MATERIAL(31).conduct = 0.173
MATERIAL(31).dens = 641
MATERIAL(31).spheat = 837
MATERIAL(31).emit = 0.9
MATERIAL(31).solAbs = 0.65
MATERIAL(31).visAbs = 0.65
MATERIAL(32).nm = "HF-C15"
MATERIAL(32).desc = "CONCRETE LW 6IN"
MATERIAL(32).rough = 1
MATERIAL(32).thick = 0.1524
MATERIAL(32).conduct = 0.173
MATERIAL(32).dens = 641
MATERIAL(32).spheat = 837
MATERIAL(32).emit = 0.9
MATERIAL(32).solAbs = 0.65
MATERIAL(32).visAbs = 0.65
MATERIAL(33).nm = "HF-C16"
MATERIAL(33).desc = "CONCRETE LW 8IN"
MATERIAL(33).rough = 1
MATERIAL(33).thick = 0.2032
MATERIAL(33).conduct = 0.173
MATERIAL(33).dens = 641
MATERIAL(33).spheat = 837
MATERIAL(33).emit = 0.9
MATERIAL(33).solAbs = 0.65
MATERIAL(33).visAbs = 0.65
gypForInterior = 34
MATERIAL(34).nm = "HF-E1"
MATERIAL(34).desc = "0.75IN PLAS - 0.75IN GYPS"
MATERIAL(34).rough = 1
MATERIAL(34).thick = 0.0191
MATERIAL(34).conduct = 0.7264
MATERIAL(34).dens = 1602
MATERIAL(34).spheat = 837
MATERIAL(34).emit = 0.9
MATERIAL(34).solAbs = 0.92
MATERIAL(34).visAbs = 0.92
MATERIAL(35).nm = "HF-E2"
MATERIAL(35).desc = "SLAG OR STONE 0.5IN"
MATERIAL(35).rough = 1
MATERIAL(35).thick = 0.0127
MATERIAL(35).conduct = 1.436
MATERIAL(35).dens = 881
MATERIAL(35).spheat = 1674
MATERIAL(35).emit = 0.9
MATERIAL(35).solAbs = 0.55
MATERIAL(35).visAbs = 0.55
MATERIAL(36).nm = "HF-E3"
MATERIAL(36).desc = "FELT AND MEMB. 0.75IN"
MATERIAL(36).rough = 1
MATERIAL(36).thick = 0.0095
MATERIAL(36).conduct = 0.1903
MATERIAL(36).dens = 1121
MATERIAL(36).spheat = 1674
MATERIAL(36).emit = 0.9
MATERIAL(36).solAbs = 0.75
MATERIAL(36).visAbs = 0.75


MATERIAL(38).nm = "HF-E5"
MATERIAL(38).desc = "ACOUSTIC TILE"
MATERIAL(38).rough = 1
MATERIAL(38).thick = 0.0191
MATERIAL(38).conduct = 0.0605
MATERIAL(38).dens = 481
MATERIAL(38).spheat = 837
MATERIAL(38).emit = 0.9
MATERIAL(38).solAbs = 0.32
MATERIAL(38).visAbs = 0.32

'initialize the window gas and glass array
windowGlassGas(1).nm = "CLEAR 2.5MM"
windowGlassGas(2).nm = "CLEAR 3MM"
windowGlassGas(3).nm = "CLEAR 6MM"
windowGlassGas(4).nm = "CLEAR 12MM"
windowGlassGas(5).nm = "BRONZE 3MM"
windowGlassGas(6).nm = "BRONZE 6MM"
windowGlassGas(7).nm = "BRONZE 10MM"
windowGlassGas(8).nm = "GREY 3MM"
windowGlassGas(9).nm = "GREY 6MM"
windowGlassGas(10).nm = "GREY 12MM"
windowGlassGas(11).nm = "GREEN 3MM"
windowGlassGas(12).nm = "GREEN 6MM"
windowGlassGas(13).nm = "LOW IRON 2.5MM"
windowGlassGas(14).nm = "LOW IRON 3MM"
windowGlassGas(15).nm = "LOW IRON 4MM"
windowGlassGas(16).nm = "LOW IRON 5MM"
windowGlassGas(17).nm = "BLUE 6MM"
windowGlassGas(18).nm = "REF A CLEAR LO 6MM"
windowGlassGas(19).nm = "REF A CLEAR MID 6MM"
windowGlassGas(20).nm = "REF A CLEAR HI 6MM"
windowGlassGas(21).nm = "REF A TINT LO 6MM"
windowGlassGas(22).nm = "REF A TINT MID 6MM"
windowGlassGas(23).nm = "REF A TINT HI 6MM"
windowGlassGas(24).nm = "REF B CLEAR LO 6MM"
windowGlassGas(25).nm = "REF B CLEAR HI 6MM"
windowGlassGas(26).nm = "REF B TINT LO 6MM"
windowGlassGas(27).nm = "REF B TINT MID 6MM"
windowGlassGas(28).nm = "REF B TINT HI 6MM"
windowGlassGas(29).nm = "REF C CLEAR LO 6MM"
windowGlassGas(30).nm = "REF C CLEAR MID 6MM"
windowGlassGas(31).nm = "REF C CLEAR HI 6MM"
windowGlassGas(32).nm = "REF C TINT LO 6MM"
windowGlassGas(33).nm = "REF C TINT MID 6MM"
windowGlassGas(34).nm = "REF C TINT HI 6MM"
windowGlassGas(35).nm = "REF D CLEAR 6MM"
windowGlassGas(36).nm = "REF D TINT 6MM"
windowGlassGas(37).nm = "PYR A CLEAR 3MM"
windowGlassGas(38).nm = "PYR B CLEAR 3MM"
windowGlassGas(39).nm = "PYR B CLEAR 6MM"
windowGlassGas(40).nm = "LoE CLEAR 3MM"
windowGlassGas(41).nm = "LoE CLEAR 3MM Rev"
windowGlassGas(42).nm = "LoE CLEAR 6MM"
windowGlassGas(43).nm = "LoE CLEAR 6MM Rev"
windowGlassGas(44).nm = "LoE TINT 6MM"
windowGlassGas(45).nm = "LoE SPEC SEL CLEAR 3MM"
windowGlassGas(46).nm = "LoE SPEC SEL CLEAR 6MM"
windowGlassGas(47).nm = "LoE SPEC SEL CLEAR 6MM Rev"
windowGlassGas(48).nm = "LoE SPEC SEL TINT 6MM"
windowGlassGas(49).nm = "COATED POLY-88"
windowGlassGas(50).nm = "COATED POLY-77"
windowGlassGas(51).nm = "COATED POLY-66"
windowGlassGas(52).nm = "COATED POLY-55"
windowGlassGas(53).nm = "COATED POLY-44"
windowGlassGas(54).nm = "COATED POLY-33"
windowGlassGas(55).nm = "ECABS-1 BLEACHED 6MM"
windowGlassGas(56).nm = "ECABS-1 COLORED 6MM"
windowGlassGas(57).nm = "ECREF-1 BLEACHED 6MM"
windowGlassGas(58).nm = "ECREF-1 COLORED 6MM"
windowGlassGas(59).nm = "ECABS-2 BLEACHED 6MM"
windowGlassGas(60).nm = "ECABS-2 COLORED 6MM"
windowGlassGas(61).nm = "ECREF-2 BLEACHED 6MM"
windowGlassGas(62).nm = "ECREF-2 COLORED 6MM"
windowGlassGas(63).nm = "AIR 3MM"
windowGlassGas(64).nm = "AIR 6MM"
windowGlassGas(65).nm = "AIR 8MM"
windowGlassGas(66).nm = "AIR 13MM"
windowGlassGas(67).nm = "ARGON 3MM"
windowGlassGas(68).nm = "ARGON 6MM"
windowGlassGas(69).nm = "ARGON 8MM"
windowGlassGas(70).nm = "ARGON 13MM"
windowGlassGas(71).nm = "KRYPTON 3MM"
windowGlassGas(72).nm = "KRYPTON 6MM"
windowGlassGas(73).nm = "KRYPTON 8MM"
windowGlassGas(74).nm = "KRYPTON 13MM"
windowGlassGas(75).nm = "XENON 3MM"
windowGlassGas(76).nm = "XENON 6MM"
windowGlassGas(77).nm = "XENON 8MM"
windowGlassGas(78).nm = "XENON 13MM"

windowGlassGas(1).prop = "MATERIAL:WindowGlass,CLEAR 2.5MM,SpectralAverage,              ,.0025,.850,.075,.075,.901,.081,.081,.0,.84,.84,.9;  ! ID 1"
windowGlassGas(2).prop = "MATERIAL:WindowGlass,CLEAR 3MM,SpectralAverage,                 ,.003,.837,.075,.075,.898,.081,.081,.0,.84,.84,.9;  ! ID 2"
windowGlassGas(3).prop = "MATERIAL:WindowGlass,CLEAR 6MM,SpectralAverage,                 ,.006,.775,.071,.071,.881,.080,.080,.0,.84,.84,.9;  ! ID 3"
windowGlassGas(4).prop = "MATERIAL:WindowGlass,CLEAR 12MM,SpectralAverage,                ,.012,.653,.064,.064,.841,.077,.077,.0,.84,.84,.9;  ! ID 4"
windowGlassGas(5).prop = "MATERIAL:WindowGlass,BRONZE 3MM,SpectralAverage,                ,.003,.645,.062,.062,.685,.065,.065,.0,.84,.84,.9;  ! ID 5"
windowGlassGas(6).prop = "MATERIAL:WindowGlass,BRONZE 6MM,SpectralAverage,                ,.006,.482,.054,.054,.534,.057,.057,.0,.84,.84,.9;  ! ID 6"
windowGlassGas(7).prop = "MATERIAL:WindowGlass,BRONZE 10MM,SpectralAverage,               ,.010,.326,.048,.048,.379,.050,.050,.0,.84,.84,.9;  ! ID 7"
windowGlassGas(8).prop = "MATERIAL:WindowGlass,GREY 3MM,SpectralAverage,                  ,.003,.626,.061,.061,.611,.061,.061,.0,.84,.84,.9;  ! ID 8"
windowGlassGas(9).prop = "MATERIAL:WindowGlass,GREY 6MM,SpectralAverage,                  ,.006,.455,.053,.053,.431,.052,.052,.0,.84,.84,.9;  ! ID 9"
windowGlassGas(10).prop = "MATERIAL:WindowGlass,GREY 12MM,SpectralAverage,                 ,.012,.217,.044,.044,.187,.045,.045,.0,.84,.84,.9;  ! ID 10"
windowGlassGas(11).prop = "MATERIAL:WindowGlass,GREEN 3MM,SpectralAverage,                 ,.003,.635,.063,.063,.822,.075,.075,.0,.84,.84,.9;  ! ID 11"
windowGlassGas(12).prop = "MATERIAL:WindowGlass,GREEN 6MM,SpectralAverage,                 ,.006,.487,.056,.056,.749,.070,.070,.0,.84,.84,.9;  ! ID 12"
windowGlassGas(13).prop = "MATERIAL:WindowGlass,LOW IRON 2.5MM,SpectralAverage,           ,.0025,.904,.080,.080,.914,.083,.083,.0,.84,.84,.9;  ! ID 13"
windowGlassGas(14).prop = "MATERIAL:WindowGlass,LOW IRON 3MM,SpectralAverage,              ,.003,.899,.079,.079,.913,.082,.082,.0,.84,.84,.9;  ! ID 14"
windowGlassGas(15).prop = "MATERIAL:WindowGlass,LOW IRON 4MM,SpectralAverage,              ,.004,.894,.079,.079,.911,.082,.082,.0,.84,.84,.9;  ! ID 15"
windowGlassGas(16).prop = "MATERIAL:WindowGlass,LOW IRON 5MM,SpectralAverage,              ,.005,.889,.079,.079,.910,.082,.082,.0,.84,.84,.9;  ! ID 16"
windowGlassGas(17).prop = "MATERIAL:WindowGlass,BLUE 6MM,SpectralAverage,                  ,.006,.480,.050,.050,.570,.060,.060,.0,.84,.84,.9;  ! ID 17"
windowGlassGas(18).prop = "MATERIAL:WindowGlass,REF A CLEAR LO 6MM,SpectralAverage,        ,.006,.066,.341,.493,.080,.410,.370,.0,.84,.40,.9;  ! ID 200"
windowGlassGas(19).prop = "MATERIAL:WindowGlass,REF A CLEAR MID 6MM,SpectralAverage,       ,.006,.110,.270,.430,.140,.310,.350,.0,.84,.47,.9;  ! ID 201"
windowGlassGas(20).prop = "MATERIAL:WindowGlass,REF A CLEAR HI 6MM,SpectralAverage,        ,.006,.159,.220,.370,.200,.250,.320,.0,.84,.57,.9;  ! ID 202"
windowGlassGas(21).prop = "MATERIAL:WindowGlass,REF A TINT LO 6MM,SpectralAverage,         ,.006,.040,.150,.470,.050,.170,.370,.0,.84,.41,.9;  ! ID 210"
windowGlassGas(22).prop = "MATERIAL:WindowGlass,REF A TINT MID 6MM,SpectralAverage,        ,.006,.060,.130,.420,.090,.140,.350,.0,.84,.47,.9;  ! ID 211"
windowGlassGas(23).prop = "MATERIAL:WindowGlass,REF A TINT HI 6MM,SpectralAverage,         ,.006,.100,.110,.380,.100,.110,.320,.0,.84,.53,.9;  ! ID 212"
windowGlassGas(24).prop = "MATERIAL:WindowGlass,REF B CLEAR LO 6MM,SpectralAverage,        ,.006,.150,.220,.380,.200,.230,.330,.0,.84,.58,.9;  ! ID 220"
windowGlassGas(25).prop = "MATERIAL:WindowGlass,REF B CLEAR HI 6MM,SpectralAverage,        ,.006,.240,.160,.320,.300,.160,.290,.0,.84,.60,.9;  ! ID 221"
windowGlassGas(26).prop = "MATERIAL:WindowGlass,REF B TINT LO 6MM,SpectralAverage,         ,.006,.040,.130,.420,.050,.090,.280,.0,.84,.41,.9;  ! ID 230"
windowGlassGas(27).prop = "MATERIAL:WindowGlass,REF B TINT MID 6MM,SpectralAverage,        ,.006,.100,.110,.410,.130,.100,.320,.0,.84,.45,.9;  ! ID 231"
windowGlassGas(28).prop = "MATERIAL:WindowGlass,REF B TINT HI 6MM,SpectralAverage,         ,.006,.150,.090,.330,.180,.080,.280,.0,.84,.60,.9;  ! ID 232"
windowGlassGas(29).prop = "MATERIAL:WindowGlass,REF C CLEAR LO 6MM,SpectralAverage,        ,.006,.110,.250,.490,.130,.280,.420,.0,.84,.43,.9;  ! ID 240"
windowGlassGas(30).prop = "MATERIAL:WindowGlass,REF C CLEAR MID 6MM,SpectralAverage,       ,.006,.170,.200,.420,.190,.210,.380,.0,.84,.51,.9;  ! ID 241"
windowGlassGas(31).prop = "MATERIAL:WindowGlass,REF C CLEAR HI 6MM,SpectralAverage,        ,.006,.200,.160,.390,.220,.170,.350,.0,.84,.55,.9;  ! ID 242"
windowGlassGas(32).prop = "MATERIAL:WindowGlass,REF C TINT LO 6MM,SpectralAverage,         ,.006,.070,.130,.490,.080,.130,.420,.0,.84,.43,.9;  ! ID 250"
windowGlassGas(33).prop = "MATERIAL:WindowGlass,REF C TINT MID 6MM,SpectralAverage,        ,.006,.100,.100,.420,.110,.100,.380,.0,.84,.51,.9;  ! ID 251"
windowGlassGas(34).prop = "MATERIAL:WindowGlass,REF C TINT HI 6MM,SpectralAverage,         ,.006,.120,.090,.390,.130,.090,.350,.0,.84,.55,.9;  ! ID 252"
windowGlassGas(35).prop = "MATERIAL:WindowGlass,REF D CLEAR 6MM,SpectralAverage,           ,.006,.429,.308,.379,.334,.453,.505,.0,.84,.82,.9;  ! ID 260"
windowGlassGas(36).prop = "MATERIAL:WindowGlass,REF D TINT 6MM,SpectralAverage,            ,.006,.300,.140,.360,.250,.180,.450,.0,.84,.82,.9;  ! ID 270"
windowGlassGas(37).prop = "MATERIAL:WindowGlass,PYR A CLEAR 3MM,SpectralAverage,           ,.003,.750,.100,.100,.850,.120,.120,.0,.84,.40,.9;  ! ID 300"
windowGlassGas(38).prop = "MATERIAL:WindowGlass,PYR B CLEAR 3MM,SpectralAverage,           ,.003,.740,.090,.100,.820,.110,.120,.0,.84,.20,.9;  ! ID 350"
windowGlassGas(39).prop = "MATERIAL:WindowGlass,PYR B CLEAR 6MM,SpectralAverage,           ,.006,.680,.090,.100,.810,.110,.120,.0,.84,.20,.9;  ! ID 351"
windowGlassGas(40).prop = "MATERIAL:WindowGlass,LoE CLEAR 3MM,SpectralAverage,             ,.003,.630,.190,.220,.850,.056,.079,.0,.84,.10,.9;  ! ID 400"
windowGlassGas(41).prop = "MATERIAL:WindowGlass,LoE CLEAR 3MM Rev,SpectralAverage,         ,.003,.630,.220,.190,.850,.079,.056,.0,.10,.84,.9;  ! "
windowGlassGas(42).prop = "MATERIAL:WindowGlass,LoE CLEAR 6MM,SpectralAverage,             ,.006,.600,.170,.220,.840,.055,.078,.0,.84,.10,.9;  ! ID 401"
windowGlassGas(43).prop = "MATERIAL:WindowGlass,LoE CLEAR 6MM Rev,SpectralAverage,         ,.006,.600,.220,.170,.840,.078,.055,.0,.10,.84,.9;  ! "
windowGlassGas(44).prop = "MATERIAL:WindowGlass,LoE TINT 6MM,SpectralAverage,              ,.006,.360,.093,.200,.500,.035,.054,.0,.84,.10,.9;  ! ID 451"
windowGlassGas(45).prop = "MATERIAL:WindowGlass,LoE SPEC SEL CLEAR 3MM,SpectralAverage,    ,.003,.450,.340,.370,.780,.070,.060,.0,.84,.03,.9;  ! ID 500"
windowGlassGas(46).prop = "MATERIAL:WindowGlass,LoE SPEC SEL CLEAR 6MM,SpectralAverage,    ,.006,.430,.300,.420,.770,.070,.060,.0,.84,.03,.9;  ! ID 501"
windowGlassGas(47).prop = "MATERIAL:WindowGlass,LoE SPEC SEL CLEAR 6MM Rev,SpectralAverage,,.006,.430,.420,.300,.770,.060,.070,.0,.03,.84,.9;  ! "
windowGlassGas(48).prop = "MATERIAL:WindowGlass,LoE SPEC SEL TINT 6MM,SpectralAverage,     ,.006,.260,.140,.410,.460,.060,.040,.0,.84,.03,.9;  ! ID 550"
windowGlassGas(49).prop = "MATERIAL:WindowGlass,COATED POLY-88,SpectralAverage,          ,.00051,.656,.249,.227,.868,.064,.060,.0,.136,.720,.14;  ! ID 600"
windowGlassGas(50).prop = "MATERIAL:WindowGlass,COATED POLY-77,SpectralAverage,          ,.00051,.504,.402,.398,.766,.147,.167,.0,.075,.720,.14;  ! ID 601"
windowGlassGas(51).prop = "MATERIAL:WindowGlass,COATED POLY-66,SpectralAverage,          ,.00051,.403,.514,.515,.658,.256,.279,.0,.057,.720,.14;  ! ID 602"
windowGlassGas(52).prop = "MATERIAL:WindowGlass,COATED POLY-55,SpectralAverage,          ,.00051,.320,.582,.593,.551,.336,.375,.0,.046,.720,.14;  ! ID 603"
windowGlassGas(53).prop = "MATERIAL:WindowGlass,COATED POLY-44,SpectralAverage,          ,.00051,.245,.626,.641,.439,.397,.453,.0,.037,.720,.14;  ! ID 604"
windowGlassGas(54).prop = "MATERIAL:WindowGlass,COATED POLY-33,SpectralAverage,          ,.00051,.178,.739,.738,.330,.566,.591,.0,.035,.720,.14;  ! ID 605"
windowGlassGas(55).prop = "MATERIAL:WindowGlass,ECABS-1 BLEACHED 6MM,SpectralAverage,      ,.006,.814,.086,.086,.847,.099,.099,.0,.84,.84,.9;  ! ID 700"
windowGlassGas(56).prop = "MATERIAL:WindowGlass,ECABS-1 COLORED 6MM,SpectralAverage,       ,.006,.111,.179,.179,.128,.081,.081,.0,.84,.84,.9;  ! ID 701"
windowGlassGas(57).prop = "MATERIAL:WindowGlass,ECREF-1 BLEACHED 6MM,SpectralAverage,      ,.006,.694,.168,.168,.818,.110,.110,.0,.84,.84,.9;  ! ID 702"
windowGlassGas(58).prop = "MATERIAL:WindowGlass,ECREF-1 COLORED 6MM,SpectralAverage,       ,.006,.099,.219,.219,.155,.073,.073,.0,.84,.84,.9;  ! ID 703"
windowGlassGas(59).prop = "MATERIAL:WindowGlass,ECABS-2 BLEACHED 6MM,SpectralAverage,      ,.006,.814,.086,.086,.847,.099,.099,.0,.84,.10,.9;  ! ID 704"
windowGlassGas(60).prop = "MATERIAL:WindowGlass,ECABS-2 COLORED 6MM,SpectralAverage,       ,.006,.111,.179,.179,.128,.081,.081,.0,.84,.10,.9;  ! ID 705"
windowGlassGas(61).prop = "MATERIAL:WindowGlass,ECREF-2 BLEACHED 6MM,SpectralAverage,      ,.006,.694,.168,.168,.818,.110,.110,.0,.84,.10,.9;  ! ID 706"
windowGlassGas(62).prop = "MATERIAL:WindowGlass,ECREF-2 COLORED 6MM,SpectralAverage,       ,.006,.099,.219,.219,.155,.073,.073,.0,.84,.10,.9;  ! ID 707"
windowGlassGas(63).prop = "MATERIAL:WindowGas, AIR 3MM, Air, .0032;"
windowGlassGas(64).prop = "MATERIAL:WindowGas, AIR 6MM, Air, .0063;"
windowGlassGas(65).prop = "MATERIAL:WindowGas, AIR 8MM, Air, .0079;    "
windowGlassGas(66).prop = "MATERIAL:WindowGas, AIR 13MM, Air, .0127;"
windowGlassGas(67).prop = "MATERIAL:WindowGas, ARGON 3MM, Argon, .0032;"
windowGlassGas(68).prop = "MATERIAL:WindowGas, ARGON 6MM, Argon, .0063;"
windowGlassGas(69).prop = "MATERIAL:WindowGas, ARGON 8MM, Argon, .0079;"
windowGlassGas(70).prop = "MATERIAL:WindowGas, ARGON 13MM, Argon, .0127;"
windowGlassGas(71).prop = "MATERIAL:WindowGas, KRYPTON 3MM, Krypton, .0032;"
windowGlassGas(72).prop = "MATERIAL:WindowGas, KRYPTON 6MM, Krypton, .0063;"
windowGlassGas(73).prop = "MATERIAL:WindowGas, KRYPTON 8MM, Krypton, .0079;"
windowGlassGas(74).prop = "MATERIAL:WindowGas, KRYPTON 13MM, Krypton, .0127;"
windowGlassGas(75).prop = "MATERIAL:WindowGas, XENON 3MM, Xenon, .0032;"
windowGlassGas(76).prop = "MATERIAL:WindowGas, XENON 6MM, Xenon, .0063;"
windowGlassGas(77).prop = "MATERIAL:WindowGas, XENON 8MM, Xenon, .0079;"
windowGlassGas(78).prop = "MATERIAL:WindowGas, XENON 13MM, Xenon, .0127;"

For iLay = 0 To 36
  windowLayers(iLay).layerCount = 1
Next iLay
For iLay = 37 To 180
  windowLayers(iLay).layerCount = 3
Next iLay
For iLay = 181 To 205
  windowLayers(iLay).layerCount = 5
Next iLay
windowLayers(206).layerCount = 7

'CONSTRUCTION Sgl Clr 3mm   ! 1000  U=6.61  SC=1.00  SHGC=.86  TSOL=.84  TVIS=.90
windowLayers(0).layerName(1) = "CLEAR 3MM"
'CONSTRUCTION Sgl Clr 6mm   ! 1001  U=6.17  SC= .95  SHGC=.81  TSOL=.77  TVIS=.88
windowLayers(1).layerName(1) = "CLEAR 6MM"
'CONSTRUCTION Sgl Clr Low Iron 3mm  ! 1002  U=6.31  SC=1.05  SHGC=.90  TSOL=.90  TVIS=.91
windowLayers(2).layerName(1) = "LOW IRON 3MM"
'CONSTRUCTION Sgl Clr Low Iron 5mm  ! 1003  U=6.22  SC=1.04  SHGC=.90  TSOL=.89  TVIS=.91
windowLayers(3).layerName(1) = "LOW IRON 5MM"
'CONSTRUCTION Sgl Bronze 3mm! 1200  U=6.31  SC= .84  SHGC=.73  TSOL=.64  TVIS=.69
windowLayers(4).layerName(1) = "BRONZE 3MM"
'CONSTRUCTION Sgl Bronze 6mm! 1201  U=6.17  SC= .71  SHGC=.61  TSOL=.48  TVIS=.53
windowLayers(5).layerName(1) = "BRONZE 6MM"
'CONSTRUCTION Sgl Green 3mm ! 1202  U=6.31  SC= .83  SHGC=.72  TSOL=.63  TVIS=.82
windowLayers(6).layerName(1) = "GREEN 3MM"
'CONSTRUCTION Sgl Green 6mm ! 1203  U=6.17  SC= .71  SHGC=.61  TSOL=.49  TVIS=.75
windowLayers(7).layerName(1) = "GREEN 6MM"
'CONSTRUCTION Sgl Grey 3mm  ! 1204  U=6.31  SC= .83  SHGC=.71  TSOL=.63  TVIS=.61
windowLayers(8).layerName(1) = "GREY 3MM"
'CONSTRUCTION Sgl Grey 6mm  ! 1205  U=6.17  SC= .69  SHGC=.59  TSOL=.46  TVIS=.43
windowLayers(9).layerName(1) = "GREY 6MM"
'CONSTRUCTION Sgl Blue 6mm  ! 1206  U=6.17  SC= .71  SHGC=.61  TSOL=.48  TVIS=.57
windowLayers(10).layerName(1) = "BLUE 6MM"
'CONSTRUCTION Sgl Ref-A-L Clr 6mm   ! 1400  U=4.90  SC= .23  SHGC=.19  TSOL=.07  TVIS=.08
windowLayers(11).layerName(1) = "REF A CLEAR LO 6MM"
'CONSTRUCTION Sgl Ref-A-M Clr 6mm   ! 1401  U=5.11  SC= .29  SHGC=.25  TSOL=.11  TVIS=.14
windowLayers(12).layerName(1) = "REF A CLEAR MID 6MM "
'CONSTRUCTION Sgl Ref-A-H Clr 6mm   ! 1402  U=5.41  SC= .36  SHGC=.31  TSOL=.16  TVIS=.20
windowLayers(13).layerName(1) = "REF A CLEAR HI 6MM"
'CONSTRUCTION Sgl Ref-A-L Tint 6mm  ! 1403  U=4.93  SC= .26  SHGC=.22  TSOL=.04  TVIS=.05
windowLayers(14).layerName(1) = "REF A TINT LO 6MM"
'CONSTRUCTION Sgl Ref-A-M Tint 6mm  ! 1404  U=5.11  SC= .29  SHGC=.25  TSOL=.06  TVIS=.09
windowLayers(15).layerName(1) = "REF A TINT MID 6MM"
'CONSTRUCTION Sgl Ref-A-H Tint 6mm  ! 1405  U=5.29  SC= .34  SHGC=.29  TSOL=.10  TVIS=.10
windowLayers(16).layerName(1) = "REF A TINT HI 6MM "
'CONSTRUCTION Sgl Ref-B-L Clr 6mm   ! 1406  U=5.44  SC= .35  SHGC=.31  TSOL=.15  TVIS=.20
windowLayers(17).layerName(1) = "REF B CLEAR LO 6MM"
'CONSTRUCTION Sgl Ref-B-H Clr 6mm   ! 1407  U=5.50  SC= .45  SHGC=.39  TSOL=.24  TVIS=.30
windowLayers(18).layerName(1) = "REF B CLEAR HI 6MM"
'CONSTRUCTION Sgl Ref-B-L Tint 6mm  ! 1408  U=4.93  SC= .26  SHGC=.23  TSOL=.04  TVIS=.05
windowLayers(19).layerName(1) = "REF B TINT LO 6MM"
'CONSTRUCTION Sgl Ref-B-M Tint 6mm  ! 1409  U=5.05  SC= .33  SHGC=.28  TSOL=.10  TVIS=.13
windowLayers(20).layerName(1) = "REF B TINT MID 6MM"
'CONSTRUCTION Sgl Ref-B-H Tint 6mm  ! 1410  U=5.50  SC= .40  SHGC=.34  TSOL=.15  TVIS=.18
windowLayers(21).layerName(1) = "REF B TINT HI 6MM"
'CONSTRUCTION Sgl Ref-C-L Clr 6mm   ! 1411  U=4.99  SC= .29  SHGC=.25  TSOL=.11  TVIS=.13
windowLayers(22).layerName(1) = "REF C CLEAR LO 6MM"
'CONSTRUCTION Sgl Ref-C-M Clr 6mm   ! 1412  U=5.23  SC= .37  SHGC=.32  TSOL=.17  TVIS=.19
windowLayers(23).layerName(1) = "REF C CLEAR MID 6MM"
'CONSTRUCTION Sgl Ref-C-H Clr 6mm   ! 1413  U=5.35  SC= .41  SHGC=.35  TSOL=.20  TVIS=.22
windowLayers(24).layerName(1) = "REF C CLEAR HI 6MM"
'CONSTRUCTION Sgl Ref-C-L Tint 6mm  ! 1414  U=4.99  SC= .29  SHGC=.25  TSOL=.07  TVIS=.08
windowLayers(25).layerName(1) = "REF C TINT LO 6MM"
'CONSTRUCTION Sgl Ref-C-M Tint 6mm  ! 1415  U=5.23  SC= .34  SHGC=.29  TSOL=.10  TVIS=.11
windowLayers(26).layerName(1) = "REF C TINT MID 6MM"
'CONSTRUCTION Sgl Ref-C-H Tint 6mm  ! 1416  U=5.35  SC= .37  SHGC=.31  TSOL=.12  TVIS=.13
windowLayers(27).layerName(1) = "REF C TINT HI 6MM"
'CONSTRUCTION Sgl Ref-D Clr 6mm ! 1417  U=6.12  SC= .58  SHGC=.50  TSOL=.43  TVIS=.33'
windowLayers(28).layerName(1) = "REF D CLEAR 6MM"
'CONSTRUCTION Sgl Ref-D Tint 6mm! 1418  U=6.12  SC= .58  SHGC=.50  TSOL=.43  TVIS=.33
windowLayers(29).layerName(1) = "REF D TINT 6MM"
'CONSTRUCTION Sgl LoE (e2=.4) Clr 3mm   ! 1600  U=4.99  SC= .91  SHGC=.78  TSOL=.75  TVIS=.85
windowLayers(30).layerName(1) = "PYR A CLEAR 3MM"
'CONSTRUCTION Sgl LoE (e2=.2) Clr 3mm   ! 1601  U=4.34  SC= .89  SHGC=.77  TSOL=.74  TVIS=.82
windowLayers(31).layerName(1) = "PYR B CLEAR 3MM"
'CONSTRUCTION Sgl LoE (e2=.2) Clr 6mm   ! 1602  U=4.27  SC= .84  SHGC=.72  TSOL=.68  TVIS=.81
windowLayers(32).layerName(1) = "PYR B CLEAR 6MM"
'CONSTRUCTION Sgl Elec Abs Bleached 6mm ! 1800  U=6.17  SC= .98  SHGC=.84  TSOL=.81  TVIS=.85
windowLayers(33).layerName(1) = "ECABS-1 BLEACHED 6MM"
'CONSTRUCTION Sgl Elec Abs Colored 6mm  ! 1801  U=6.17  SC= .36  SHGC=.31  TSOL=.11  TVIS=.13
windowLayers(34).layerName(1) = "ECABS-1 COLORED 6MM"
'CONSTRUCTION Sgl Elec Ref Bleached 6mm ! 1802  U=6.17  SC= .85  SHGC=.73  TSOL=.69  TVIS=.82
windowLayers(35).layerName(1) = "ECREF-1 BLEACHED 6MM"
'CONSTRUCTION Sgl Elec Ref Colored 6mm  ! 1803  U=6.17  SC= .34  SHGC=.29  TSOL=.10  TVIS=.16
windowLayers(36).layerName(1) = "ECREF-1 COLORED 6MM"
'CONSTRUCTION Dbl Clr 3mm/6mm Air   ! 2000  U=3.23  SC= .88  SHGC=.76  TSOL=.70  TVIS=.81
windowLayers(37).layerName(1) = "CLEAR 3MM"
windowLayers(37).layerName(2) = "AIR 6MM"
windowLayers(37).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Clr 3mm/13mm Air  ! 2001  U=2.79  SC= .89  SHGC=.76  TSOL=.70  TVIS=.81
windowLayers(38).layerName(1) = "CLEAR 3MM"
windowLayers(38).layerName(2) = "AIR 13MM"
windowLayers(38).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Clr 3mm/13mm Arg  ! 2002  U=2.61  SC= .89  SHGC=.76  TSOL=.70  TVIS=.81
windowLayers(39).layerName(1) = "CLEAR 3MM"
windowLayers(39).layerName(2) = "ARGON 13MM"
windowLayers(39).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Clr 6mm/6mm Air   ! 2003  U=3.16  SC= .81  SHGC=.69  TSOL=.60  TVIS=.78
windowLayers(40).layerName(1) = "CLEAR 6MM"
windowLayers(40).layerName(2) = "AIR 6MM"
windowLayers(40).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Clr 6mm/13mm Air  ! 2004  U=2.74  SC= .81  SHGC=.70  TSOL=.60  TVIS=.78
windowLayers(41).layerName(1) = "CLEAR 6MM"
windowLayers(41).layerName(2) = "AIR 13MM"
windowLayers(41).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Clr 6mm/13mm Arg  ! 2005  U=2.56  SC= .81  SHGC=.70  TSOL=.60  TVIS=.78
windowLayers(42).layerName(1) = "CLEAR 6MM"
windowLayers(42).layerName(2) = "ARGON 13MM"
windowLayers(42).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Clr Low Iron 3mm/6mm Air  ! 2006  U=3.23  SC= .96  SHGC=.83  TSOL=.81  TVIS=.84
windowLayers(43).layerName(1) = "LOW IRON 3MM"
windowLayers(43).layerName(2) = "AIR 6MM"
windowLayers(43).layerName(3) = "LOW IRON 3MM"
'CONSTRUCTION Dbl Clr Low Iron 3mm/13mm Air ! 2007  U=2.79  SC= .96  SHGC=.83  TSOL=.81  TVIS=.84
windowLayers(44).layerName(1) = "LOW IRON 3MM"
windowLayers(44).layerName(2) = "AIR 13MM"
windowLayers(44).layerName(3) = "LOW IRON 3MM"
'CONSTRUCTION Dbl Clr Low Iron 3mm/13mm Arg ! 2008  U=2.61  SC= .96  SHGC=.83  TSOL=.81  TVIS=.84
windowLayers(45).layerName(1) = "LOW IRON 3MM"
windowLayers(45).layerName(2) = "ARGON 13MM"
windowLayers(45).layerName(3) = "LOW IRON 3MM"
'CONSTRUCTION Dbl Clr Low Iron 5mm/6mm Air  ! 2009  U=3.18  SC= .95  SHGC=.82  TSOL=.80  TVIS=.83
windowLayers(46).layerName(1) = "LOW IRON 5MM"
windowLayers(46).layerName(2) = "AIR 6MM"
windowLayers(46).layerName(3) = "LOW IRON 5MM"
'CONSTRUCTION Dbl Clr Low Iron 5mm/13mm Air ! 2010  U=2.76  SC= .95  SHGC=.82  TSOL=.80  TVIS=.83
windowLayers(47).layerName(1) = "LOW IRON 5MM"
windowLayers(47).layerName(2) = "AIR 13MM"
windowLayers(47).layerName(3) = "LOW IRON 5MM"
'CONSTRUCTION Dbl Clr Low Iron 5mm/13mm Arg ! 2011  U=2.58  SC= .96  SHGC=.83  TSOL=.81  TVIS=.84
windowLayers(48).layerName(1) = "LOW IRON 5MM"
windowLayers(48).layerName(2) = "ARGON 13MM"
windowLayers(48).layerName(3) = "LOW IRON 5MM"
'CONSTRUCTION Dbl Bronze 3mm/6mm Air! 2200  U=3.23  SC= .72  SHGC=.62  TSOL=.54  TVIS=.62
windowLayers(49).layerName(1) = "BRONZE 3MM"
windowLayers(49).layerName(2) = "AIR 6MM"
windowLayers(49).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Bronze 3mm/13mm Air   ! 2201  U=2.79  SC= .72  SHGC=.62  TSOL=.54  TVIS=.62
windowLayers(50).layerName(1) = "BRONZE 3MM"
windowLayers(50).layerName(2) = "AIR 13MM"
windowLayers(50).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Bronze 3mm/13mm Arg   ! 2202  U=2.61  SC= .72  SHGC=.62  TSOL=.54  TVIS=.62
windowLayers(51).layerName(1) = "BRONZE 3MM"
windowLayers(51).layerName(2) = "ARGON 13MM"
windowLayers(51).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Bronze 6mm/6mm Air! 2203  U=3.16  SC= .57  SHGC=.49  TSOL=.38  TVIS=.47
windowLayers(52).layerName(1) = "BRONZE 6MM"
windowLayers(52).layerName(2) = "AIR 6MM"
windowLayers(52).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Bronze 6mm/13mm Air   ! 2204  U=2.74  SC= .57  SHGC=.49  TSOL=.38  TVIS=.47
windowLayers(53).layerName(1) = "BRONZE 6MM"
windowLayers(53).layerName(2) = "AIR 13MM"
windowLayers(53).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Bronze 6mm/13mm Arg   ! 2205  U=2.56  SC= .56  SHGC=.49  TSOL=.38  TVIS=.47
windowLayers(54).layerName(1) = "BRONZE 6MM"
windowLayers(54).layerName(2) = "ARGON 13MM"
windowLayers(54).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Green 3m/6mm Air  ! 2206  U=3.23  SC= .72  SHGC=.62  TSOL=.53  TVIS=.74
windowLayers(55).layerName(1) = "GREEN 3MM"
windowLayers(55).layerName(2) = "AIR 6MM"
windowLayers(55).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Green 3mm/13mm Air! 2207  U=2.79  SC= .71  SHGC=.61  TSOL=.53  TVIS=.74
windowLayers(56).layerName(1) = "GREEN 3MM"
windowLayers(56).layerName(2) = "AIR 13MM"
windowLayers(56).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Green 3mm/13mm Arg! 2208  U=2.61  SC= .71  SHGC=.61  TSOL=.53  TVIS=.74
windowLayers(57).layerName(1) = "GREEN 3MM"
windowLayers(57).layerName(2) = "ARGON 13MM"
windowLayers(57).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Green 6mm/6mm Air ! 2209  U=3.16  SC= .58  SHGC=.50  TSOL=.38  TVIS=.66
windowLayers(58).layerName(1) = "GREEN 6MM"
windowLayers(58).layerName(2) = "AIR 6MM"
windowLayers(58).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Green 6mm/13mm Air! 2210  U=2.74  SC= .57  SHGC=.49  TSOL=.38  TVIS=.66
windowLayers(59).layerName(1) = "GREEN 6MM"
windowLayers(59).layerName(2) = "AIR 13MM"
windowLayers(59).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Green 6mm/13mm Arg! 2211  U=2.56  SC= .57  SHGC=.49  TSOL=.38  TVIS=.66
windowLayers(60).layerName(1) = "GREEN 6MM"
windowLayers(60).layerName(2) = "ARGON 13MM"
windowLayers(60).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Grey 3mm/6mm Air  ! 2212  U=3.23  SC= .71  SHGC=.61  TSOL=.53  TVIS=.55
windowLayers(61).layerName(1) = "GREY 3MM"
windowLayers(61).layerName(2) = "AIR 6MM"
windowLayers(61).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Grey 3mm/13mm Air ! 2213  U=2.79  SC= .71  SHGC=.61  TSOL=.53  TVIS=.55
windowLayers(62).layerName(1) = "GREY 3MM"
windowLayers(62).layerName(2) = "AIR 13MM"
windowLayers(62).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Grey 3mm/13mm Arg ! 2214  U=2.61  SC= .70  SHGC=.61  TSOL=.53  TVIS=.55
windowLayers(63).layerName(1) = "GREY 3MM"
windowLayers(63).layerName(2) = "ARGON 13MM"
windowLayers(63).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl Grey 6mm/6mm Air  ! 2215  U=3.16  SC= .55  SHGC=.47  TSOL=.35  TVIS=.38
windowLayers(64).layerName(1) = "GREY 6MM"
windowLayers(64).layerName(2) = "AIR 6MM"
windowLayers(64).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Grey 6mm/13mm Air ! 2216  U=2.74  SC= .54  SHGC=.47  TSOL=.35  TVIS=.38
windowLayers(65).layerName(1) = "GREY 6MM"
windowLayers(65).layerName(2) = "AIR 13MM"
windowLayers(65).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Grey 6mm/13mm Arg ! 2217  U=2.56  SC= .54  SHGC=.47  TSOL=.35  TVIS=.38
windowLayers(66).layerName(1) = "GREY 6MM"
windowLayers(66).layerName(2) = "ARGON 13MM"
windowLayers(66).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Blue 6mm/6mm Air  ! 2218  U=3.16  SC= .57  SHGC=.49  TSOL=.37  TVIS=.50
windowLayers(67).layerName(1) = "BLUE 6MM"
windowLayers(67).layerName(2) = "AIR 6MM"
windowLayers(67).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Blue 6mm/13mm Air ! 2219  U=2.74  SC= .57  SHGC=.49  TSOL=.37  TVIS=.50
windowLayers(68).layerName(1) = "BLUE 6MM"
windowLayers(68).layerName(2) = "AIR 13MM"
windowLayers(68).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Blue 6mm/13mm Arg ! 2220  U=2.56  SC= .56  SHGC=.49  TSOL=.37  TVIS=.50
windowLayers(69).layerName(1) = "BLUE 6MM"
windowLayers(69).layerName(2) = "ARGON 13MM"
windowLayers(69).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-L Clr 6mm/6mm Air   ! 2400  U=2.79  SC= .17  SHGC=.14  TSOL=.05  TVIS=.07
windowLayers(70).layerName(1) = "REF A CLEAR LO 6MM"
windowLayers(70).layerName(2) = "AIR 6MM"
windowLayers(70).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-L Clr 6mm/13mm Air  ! 2401  U=2.26  SC= .15  SHGC=.13  TSOL=.05  TVIS=.07
windowLayers(71).layerName(1) = "REF A CLEAR LO 6MM"
windowLayers(71).layerName(2) = "AIR 13MM"
windowLayers(71).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-L Clr 6mm/13mm Arg  ! 2402  U=2.02  SC= .14  SHGC=.12  TSOL=.05  TVIS=.07
windowLayers(72).layerName(1) = "REF A CLEAR LO 6MM"
windowLayers(72).layerName(2) = "ARGON 13MM"
windowLayers(72).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-M Clr 6mm/6mm Air   ! 2403  U=2.86  SC= .22  SHGC=.19  TSOL=.09  TVIS=.13
windowLayers(73).layerName(1) = "REF A CLEAR MID 6MM"
windowLayers(73).layerName(2) = "AIR 6MM"
windowLayers(73).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-M Clr 6mm/13mm Air  ! 2404  U=2.35  SC= .20  SHGC=.17  TSOL=.09  TVIS=.13
windowLayers(74).layerName(1) = "REF A CLEAR MID 6MM"
windowLayers(74).layerName(2) = "AIR 13MM"
windowLayers(74).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-M Clr 6mm/13mm Arg  ! 2405  U=2.13  SC= .20  SHGC=.17  TSOL=.09  TVIS=.13
windowLayers(75).layerName(1) = "REF A CLEAR MID 6MM"
windowLayers(75).layerName(2) = "ARGON 13MM"
windowLayers(75).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-H Clr 6mm/6mm Air   ! 2406  U=2.95  SC= .27  SHGC=.23  TSOL=.13  TVIS=.18
windowLayers(76).layerName(1) = "REF A CLEAR HI 6MM"
windowLayers(76).layerName(2) = "AIR 6MM"
windowLayers(76).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-H 6mm/13mm Air  ! 2407  U=2.47  SC= .26  SHGC=.22  TSOL=.13  TVIS=.18
windowLayers(77).layerName(1) = "REF A CLEAR HI 6MM"
windowLayers(77).layerName(2) = "AIR 13MM"
windowLayers(77).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-H Clr 6mm/13mm Arg  ! 2408  U=2.26  SC= .25  SHGC=.22  TSOL=.13  TVIS=.18
windowLayers(78).layerName(1) = "REF A CLEAR HI 6MM"
windowLayers(78).layerName(2) = "ARGON 13MM"
windowLayers(78).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-L Tint 6mm/6mm Air  ! 2410  U=2.80  SC= .18  SHGC=.15  TSOL=.03  TVIS=.05
windowLayers(79).layerName(1) = "REF A TINT LO 6MM"
windowLayers(79).layerName(2) = "AIR 6MM"
windowLayers(79).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-L Tint 6mm/13mm Air ! 2411  U=2.27  SC= .15  SHGC=.13  TSOL=.03  TVIS=.05
windowLayers(80).layerName(1) = "REF A TINT LO 6MM"
windowLayers(80).layerName(2) = "AIR 13MM"
windowLayers(80).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-L Tint 6mm/13mm Arg ! 2412  U=2.04  SC= .15  SHGC=.13  TSOL=.03  TVIS=.05
windowLayers(81).layerName(1) = "REF A TINT LO 6MM"
windowLayers(81).layerName(2) = "ARGON 13MM"
windowLayers(81).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-M Tint 6mm/6mm Air  ! 2413  U=2.86  SC= .20  SHGC=.17  TSOL=.05  TVIS=.08
windowLayers(82).layerName(1) = "REF A TINT MID 6MM"
windowLayers(82).layerName(2) = "AIR 6MM"
windowLayers(82).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-M Tint 6mm/13mm Air ! 2414  U=2.35  SC= .18  SHGC=.15  TSOL=.05  TVIS=.08
windowLayers(83).layerName(1) = "REF A TINT MID 6MM"
windowLayers(83).layerName(2) = "AIR 13MM"
windowLayers(83).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-M Tint 6mm/13mm Arg ! 2415  U=2.13  SC= .17  SHGC=.15  TSOL=.05  TVIS=.08
windowLayers(84).layerName(1) = "REF A TINT MID 6MM"
windowLayers(84).layerName(2) = "ARGON 13MM"
windowLayers(84).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-H Tint 6mm/6mm Air  ! 2416  U=2.92  SC= .24  SHGC=.21  TSOL=.08  TVIS=.09
windowLayers(85).layerName(1) = "REF A TINT HI 6MM"
windowLayers(85).layerName(2) = "AIR 6MM"
windowLayers(85).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-H Tint 6mm/13mm Air ! 2417  U=2.42  SC= .22  SHGC=.19  TSOL=.08  TVIS=.09
windowLayers(86).layerName(1) = "REF A TINT HI 6MM"
windowLayers(86).layerName(2) = "AIR 13MM"
windowLayers(86).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-A-H Tint 6mm/13mm Arg ! 2418  U=2.21  SC= .21  SHGC=.19  TSOL=.08  TVIS=.09
windowLayers(87).layerName(1) = "REF A TINT HI 6MM"
windowLayers(87).layerName(2) = "ARGON 13MM"
windowLayers(87).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-L Clr 6mm/6mm Air   ! 2420  U=2.96  SC= .27  SHGC=.23  TSOL=.12  TVIS=.18
windowLayers(88).layerName(1) = "REF B CLEAR LO 6MM"
windowLayers(88).layerName(2) = "AIR 6MM"
windowLayers(88).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-L Clr 6mm/13mm Air  ! 2421  U=2.48  SC= .25  SHGC=.22  TSOL=.12  TVIS=.18
windowLayers(89).layerName(1) = "REF B CLEAR LO 6MM"
windowLayers(89).layerName(2) = "AIR 13MM"
windowLayers(89).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-L Clr 6mm/13mm Arg  ! 2422  U=2.27  SC= .25  SHGC=.21  TSOL=.12  TVIS=.18
windowLayers(90).layerName(1) = "REF B CLEAR LO 6MM"
windowLayers(90).layerName(2) = "ARGON 13MM"
windowLayers(90).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-H Clr 6mm/6mm Air   ! 2426  U=2.98  SC= .35  SHGC=.30  TSOL=.19  TVIS=.27
windowLayers(91).layerName(1) = "REF B CLEAR HI 6MM"
windowLayers(91).layerName(2) = "AIR 6MM"
windowLayers(91).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-H Clr 6mm/13mm Air  ! 2427  U=2.50  SC= .34  SHGC=.29  TSOL=.19  TVIS=.27
windowLayers(92).layerName(1) = "REF B CLEAR HI 6MM"
windowLayers(92).layerName(2) = "AIR 13MM"
windowLayers(92).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-H Clr 6mm/13mm Arg  ! 2428  U=2.30  SC= .34  SHGC=.29  TSOL=.19  TVIS=.27
windowLayers(93).layerName(1) = "REF B CLEAR HI 6MM"
windowLayers(93).layerName(2) = "ARGON 13MM"
windowLayers(93).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-L Tint 6mm/6mm Air  ! 2430  U=2.80  SC= .18  SHGC=.15  TSOL=.03  TVIS=.05
windowLayers(94).layerName(1) = "REF B TINT LO 6MM"
windowLayers(94).layerName(2) = "AIR 6MM"
windowLayers(94).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-L Tint 6mm/13mm Air ! 2431  U=2.27  SC= .16  SHGC=.14  TSOL=.03  TVIS=.05
windowLayers(95).layerName(1) = "REF B TINT LO 6MM"
windowLayers(95).layerName(2) = "AIR 13MM"
windowLayers(95).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-L Tint 6mm/13mm Arg ! 2432  U=2.04  SC= .15  SHGC=.13  TSOL=.03  TVIS=.05
windowLayers(96).layerName(1) = "REF B TINT LO 6MM"
windowLayers(96).layerName(2) = "ARGON 13MM"
windowLayers(96).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-M Tint 6mm/6mm Air  ! 2433  U=2.84  SC= .24  SHGC=.20  TSOL=.08  TVIS=.12
windowLayers(97).layerName(1) = "REF B TINT MID 6MM"
windowLayers(97).layerName(2) = "AIR 6MM"
windowLayers(97).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-M Tint 6mm/13mm Air ! 2434  U=2.33  SC= .22  SHGC=.19  TSOL=.08  TVIS=.12
windowLayers(98).layerName(1) = "REF B TINT MID 6MM"
windowLayers(98).layerName(2) = "AIR 13MM"
windowLayers(98).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-M Tint 6mm/13mm Arg ! 2435  U=2.10  SC= .21  SHGC=.18  TSOL=.08  TVIS=.12
windowLayers(99).layerName(1) = "REF B TINT MID 6MM"
windowLayers(99).layerName(2) = "ARGON 13MM"
windowLayers(99).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-H Tint 6mm/6mm Air  ! 2436  U=2.98  SC= .29  SHGC=.25  TSOL=.12  TVIS=.16
windowLayers(100).layerName(1) = "REF B TINT HI 6MM"
windowLayers(100).layerName(2) = "AIR 6MM"
windowLayers(100).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-B-H Tint 6mm/13mm Air ! 2437  U=2.50  SC= .27  SHGC=.23  TSOL=.12  TVIS=.16
windowLayers(101).layerName(1) = "REF B TINT HI 6MM"
windowLayers(101).layerName(2) = "AIR 13MM"
windowLayers(101).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref B-H Tint 6mm/13mm Arg ! 2438  U=2.30  SC= .27  SHGC=.23  TSOL=.12  TVIS=.16
windowLayers(102).layerName(1) = "REF B TINT HI 6MM"
windowLayers(102).layerName(2) = "ARGON 13MM"
windowLayers(102).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-L Clr 6mm/6mm Air   ! 2440  U=2.82  SC= .22  SHGC=.19  TSOL=.09  TVIS=.12
windowLayers(103).layerName(1) = "REF C CLEAR LO 6MM"
windowLayers(103).layerName(2) = "AIR 6MM"
windowLayers(103).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-L Clr 6mm/13mm Air  ! 2441  U=2.30  SC= .20  SHGC=.18  TSOL=.09  TVIS=.12
windowLayers(104).layerName(1) = "REF C CLEAR LO 6MM"
windowLayers(104).layerName(2) = "AIR 13MM"
windowLayers(104).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-L Clr 6mm/13mm Arg  ! 2442  U=2.07  SC= .20  SHGC=.17  TSOL=.09  TVIS=.12
windowLayers(105).layerName(1) = "REF C CLEAR LO 6MM"
windowLayers(105).layerName(2) = "ARGON 13MM"
windowLayers(105).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-M Clr 6mm/6mm Air   ! 2443  U=2.90  SC= .28  SHGC=.24  TSOL=.14  TVIS=.17
windowLayers(106).layerName(1) = "REF C CLEAR MID 6MM"
windowLayers(106).layerName(2) = "AIR 6MM"
windowLayers(106).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-M Clr 6mm/13mm Air  ! 2444  U=2.40  SC= .27  SHGC=.23  TSOL=.14  TVIS=.17
windowLayers(107).layerName(1) = "REF C CLEAR MID 6MM"
windowLayers(107).layerName(2) = "AIR 13MM"
windowLayers(107).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-M Clr 6mm/13mm Arg  ! 2445  U=2.18  SC= .26  SHGC=.23  TSOL=.14  TVIS=.17
windowLayers(108).layerName(1) = "REF C CLEAR MID 6MM"
windowLayers(108).layerName(2) = "ARGON 13MM"
windowLayers(108).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-H Clr 6mm/6mm Air   ! 2446  U=2.94  SC= .32  SHGC=.27  TSOL=.16  TVIS=.20
windowLayers(109).layerName(1) = "REF C CLEAR HI 6MM"
windowLayers(109).layerName(2) = "AIR 6MM"
windowLayers(109).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-H Clr 6mm/13mm Air  ! 2447  U=2.45  SC= .30  SHGC=.26  TSOL=.16  TVIS=.20
windowLayers(110).layerName(1) = "REF C CLEAR HI 6MM"
windowLayers(110).layerName(2) = "AIR 13MM"
windowLayers(110).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-H Clr 6mm/13mm Arg  ! 2448  U=2.23  SC= .30  SHGC=.26  TSOL=.16  TVIS=.20
windowLayers(111).layerName(1) = "REF C CLEAR HI 6MM"
windowLayers(111).layerName(2) = "ARGON 13MM"
windowLayers(111).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-L Tint 6mm/6mm Air  ! 2450  U=2.82  SC= .21  SHGC=.18  TSOL=.06  TVIS=.07
windowLayers(112).layerName(1) = "REF C TINT LO 6MM"
windowLayers(112).layerName(2) = "AIR 6MM"
windowLayers(112).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-L Tint 6mm/13mm Air ! 2451  U=2.30  SC= .19  SHGC=.16  TSOL=.06  TVIS=.07
windowLayers(113).layerName(1) = "REF C TINT LO 6MM"
windowLayers(113).layerName(2) = "AIR 13MM"
windowLayers(113).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-L Tint 6mm/13mm Arg ! 2452  U=2.07  SC= .18  SHGC=.15  TSOL=.06  TVIS=.07
windowLayers(114).layerName(1) = "REF C TINT LO 6MM"
windowLayers(114).layerName(2) = "ARGON 13MM"
windowLayers(114).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-M Tint 6mm/6mm Air  ! 2453  U=2.90  SC= .24  SHGC=.21  TSOL=.08  TVIS=.10
windowLayers(115).layerName(1) = "REF C TINT MID 6MM"
windowLayers(115).layerName(2) = "AIR 6MM"
windowLayers(115).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-M Tint 6mm/13mm Air ! 2454  U=2.40  SC= .22  SHGC=.19  TSOL=.08  TVIS=.10
windowLayers(116).layerName(1) = "REF C TINT MID 6MM"
windowLayers(116).layerName(2) = "AIR 13MM"
windowLayers(116).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-M Tint 6mm/13mm Arg ! 2455  U=2.18  SC= .21  SHGC=.19  TSOL=.08  TVIS=.10
windowLayers(117).layerName(1) = "REF C TINT MID 6MM"
windowLayers(117).layerName(2) = "ARGON 13MM"
windowLayers(117).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-H Tint 6mm/6mm Air  ! 2456  U=2.94  SC= .26 SHGC=.23  TSOL=.10  TVIS=.12
windowLayers(118).layerName(1) = "REF C TINT HI 6MM"
windowLayers(118).layerName(2) = "AIR 6MM"
windowLayers(118).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-H Tint 6mm/13mm Air ! 2457  U=2.45  SC= .24  SHGC=.21  TSOL=.10  TVIS=.12
windowLayers(119).layerName(1) = "REF C TINT HI 6MM"
windowLayers(119).layerName(2) = "AIR 13MM"
windowLayers(119).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-C-H Tint 6mm/13mm Arg ! 2458  U=2.23  SC= .24  SHGC=.20  TSOL=.10  TVIS=.12
windowLayers(120).layerName(1) = "REF C TINT HI 6MM"
windowLayers(120).layerName(2) = "ARGON 13MM"
windowLayers(120).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-D Clr 6mm/6mm Air ! 2460  U=3.15  SC= .49  SHGC=.42  TSOL=.34  TVIS=.31
windowLayers(121).layerName(1) = "REF D CLEAR 6MM"
windowLayers(121).layerName(2) = "AIR 6MM"
windowLayers(121).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-D Clr 6mm/13mm Air! 2461  U=2.72  SC= .49  SHGC=.42  TSOL=.34  TVIS=.31
windowLayers(122).layerName(1) = "REF D CLEAR 6MM"
windowLayers(122).layerName(2) = "AIR 13MM"
windowLayers(122).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-D Clr 6mm/13mm Arg! 2462  U=2.54  SC= .49  SHGC=.42  TSOL=.34  TVIS=.31
windowLayers(123).layerName(1) = "REF D CLEAR 6MM"
windowLayers(123).layerName(2) = "ARGON 13MM"
windowLayers(123).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-D Tint 6mm/6mm Air! 2470  U=3.15  SC= .41  SHGC=.35  TSOL=.24  TVIS=.23
windowLayers(124).layerName(1) = "REF D TINT 6MM"
windowLayers(124).layerName(2) = "AIR 6MM"
windowLayers(124).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-D Tint 6mm/13mm Air   ! 2471  U=2.72  SC= .40  SHGC=.35  TSOL=.24  TVIS=.23
windowLayers(125).layerName(1) = "REF D TINT 6MM"
windowLayers(125).layerName(2) = "AIR 13MM"
windowLayers(125).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Ref-D Tint 6mm/13mm Arg   ! 2472  U=2.54  SC= .40  SHGC=.34  TSOL=.24  TVIS=.23
windowLayers(126).layerName(1) = "REF D TINT 6MM"
windowLayers(126).layerName(2) = "ARGON 13MM"
windowLayers(126).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.4) Clr 3mm/6mm Air   ! 2600  U=2.85  SC= .84  SHGC=.72  TSOL=.63  TVIS=.77
windowLayers(127).layerName(1) = "PYR A CLEAR 3MM"
windowLayers(127).layerName(2) = "AIR 6MM"
windowLayers(127).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.4) Clr 3mm/13mm Air  ! 2601  U=2.30  SC= .85  SHGC=.73  TSOL=.63  TVIS=.77
windowLayers(128).layerName(1) = "PYR A CLEAR 3MM"
windowLayers(128).layerName(2) = "AIR 13MM"
windowLayers(128).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.4) Clr 3mm/13mm Arg  ! 2602  U=2.05  SC= .85  SHGC=.73  TSOL=.63  TVIS=.77
windowLayers(129).layerName(1) = "PYR A CLEAR 3MM"
windowLayers(129).layerName(2) = "ARGON 13MM"
windowLayers(129).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.2) Clr 3mm/6mm Air   ! 2610  U=2.61  SC= .84  SHGC=.72  TSOL=.62  TVIS=.74
windowLayers(130).layerName(1) = "PYR B CLEAR 3MM"
windowLayers(130).layerName(2) = "AIR 6MM"
windowLayers(130).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.2) Clr 3mm/13mm Air  ! 2611  U=1.99  SC= .85  SHGC=.73  TSOL=.62  TVIS=.74
windowLayers(131).layerName(1) = "PYR B CLEAR 3MM"
windowLayers(131).layerName(2) = "AIR 13MM"
windowLayers(131).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.2) Clr 3mm/13mm Arg  ! 2612  U=1.70  SC= .86  SHGC=.74  TSOL=.62  TVIS=.74
windowLayers(132).layerName(1) = "PYR B CLEAR 3MM"
windowLayers(132).layerName(2) = "ARGON 13MM"
windowLayers(132).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.2) Clr 6mm/6mm Air   ! 2613  U=2.57  SC= .77  SHGC=.66  TSOL=.53  TVIS=.72
windowLayers(133).layerName(1) = "PYR B CLEAR 6MM"
windowLayers(133).layerName(2) = "AIR 6MM"
windowLayers(133).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.2) Clr 6mm/13mm Air  ! 2614  U=1.96  SC= .78  SHGC=.67  TSOL=.53  TVIS=.72
windowLayers(134).layerName(1) = "PYR B CLEAR 6MM"
windowLayers(134).layerName(2) = "AIR 13MM"
windowLayers(134).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.2) Clr 6mm/13mm Arg  ! 2615  U=1.67  SC= .79  SHGC=.68  TSOL=.53  TVIS=.72
windowLayers(135).layerName(1) = "PYR B CLEAR 6MM"
windowLayers(135).layerName(2) = "ARGON 13MM"
windowLayers(135).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.1) Clr 3mm/6mm Air   ! 2630  U=2.47  SC= .69  SHGC=.60  TSOL=.54  TVIS=.77
windowLayers(136).layerName(1) = "LoE CLEAR 3MM"
windowLayers(136).layerName(2) = "AIR 6MM"
windowLayers(136).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.1) Clr 3mm/13mm Air  ! 2631  U=1.81  SC= .69  SHGC=.60  TSOL=.54  TVIS=.77
windowLayers(137).layerName(1) = "LoE CLEAR 3MM"
windowLayers(137).layerName(2) = "AIR 13MM"
windowLayers(137).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.1) Clr 3mm/13mm Arg  ! 2632  U=1.48  SC= .69  SHGC=.59  TSOL=.54  TVIS=.77
windowLayers(138).layerName(1) = "LoE CLEAR 3MM"
windowLayers(138).layerName(2) = "ARGON 13MM"
windowLayers(138).layerName(3) = "CLEAR 3MM"
'CONSTRUCTION Dbl LoE (e2=.1) Clr 6mm/6mm Air   ! 2633  U=2.43  SC= .65  SHGC=.56  TSOL=.47  TVIS=.75
windowLayers(139).layerName(1) = "LoE CLEAR 6MM"
windowLayers(139).layerName(2) = "AIR 6MM"
windowLayers(139).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.1) Clr 6mm/13mm Air  ! 2634  U=1.78  SC= .65  SHGC=.56  TSOL=.47  TVIS=.75
windowLayers(140).layerName(1) = "LoE CLEAR 6MM"
windowLayers(140).layerName(2) = "AIR 13MM"
windowLayers(140).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.1) Clr 6mm/13mm Arg  ! 2635  U=1.46  SC= .66  SHGC=.56  TSOL=.47  TVIS=.75
windowLayers(141).layerName(1) = "LoE CLEAR 6MM"
windowLayers(141).layerName(2) = "ARGON 13MM"
windowLayers(141).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.1) Tint 6mm/6mm Air  ! 2636  U=2.43  SC= .45  SHGC=.39  TSOL=.28  TVIS=.44
windowLayers(142).layerName(1) = "LoE TINT 6MM"
windowLayers(142).layerName(2) = "AIR 6MM"
windowLayers(142).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.1) Tint 6mm/13mm Air ! 2637  U=1.78  SC= .43  SHGC=.37  TSOL=.28  TVIS=.44
windowLayers(143).layerName(1) = "LoE TINT 6MM"
windowLayers(143).layerName(2) = "AIR 13MM"
windowLayers(143).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e2=.1) Tint 6mm/13mm Arg ! 2638  U=1.46  SC= .43  SHGC=.37  TSOL=.28  TVIS=.44
windowLayers(144).layerName(1) = "LoE TINT 6MM"
windowLayers(144).layerName(2) = "ARGON 13MM"
windowLayers(144).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE (e3=.1) Clr 3mm/6mm Air   ! 2640  U=2.61  SC= .84  SHGC=.72  TSOL=.62  TVIS=.74
windowLayers(145).layerName(1) = "CLEAR 3MM"
windowLayers(145).layerName(2) = "AIR 6MM"
windowLayers(145).layerName(3) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Dbl LoE (e3=.1) Clr 3mm/13mm Air  ! 2641  U=1.99  SC= .85  SHGC=.73  TSOL=.62  TVIS=.74
windowLayers(146).layerName(1) = "CLEAR 3MM"
windowLayers(146).layerName(2) = "AIR 13MM"
windowLayers(146).layerName(3) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Dbl LoE (e3=.1) Clr 3mm/13mm Arg  ! 2642  U=1.70  SC= .86  SHGC=.74  TSOL=.62  TVIS=.74
windowLayers(147).layerName(1) = "CLEAR 3MM"
windowLayers(147).layerName(2) = "ARGON 13MM"
windowLayers(147).layerName(3) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Dbl LoE Spec Sel Clr 3mm/6mm/6mm Air  ! 2660  U=2.38  SC= .51  SHGC=.44  TSOL=.39  TVIS=.70
windowLayers(148).layerName(1) = "LoE SPEC SEL CLEAR 3MM"
windowLayers(148).layerName(2) = "AIR 6MM"
windowLayers(148).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Clr 3mm/13mm/6mm Air ! 2661  U=1.68  SC= .51  SHGC=.44  TSOL=.39  TVIS=.70
windowLayers(149).layerName(1) = "LoE SPEC SEL CLEAR 3MM"
windowLayers(149).layerName(2) = "AIR 13MM"
windowLayers(149).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Clr 3mm/13mm/6mm Arg ! 2662  U=1.34  SC= .50  SHGC=.43  TSOL=.39  TVIS=.70
windowLayers(150).layerName(1) = "LoE SPEC SEL CLEAR 3MM"
windowLayers(150).layerName(2) = "ARGON 13MM"
windowLayers(150).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Clr 6mm/6mm Air  ! 2663  U=2.41  SC= .49  SHGC=.42  TSOL=.34  TVIS=.68
windowLayers(151).layerName(1) = "LoE SPEC SEL CLEAR 6MM"
windowLayers(151).layerName(2) = "AIR 6MM"
windowLayers(151).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Clr 6mm/13mm Air ! 2664  U=1.67  SC= .48  SHGC=.42  TSOL=.34  TVIS=.68
windowLayers(152).layerName(1) = "LoE SPEC SEL CLEAR 6MM"
windowLayers(152).layerName(2) = "AIR 13MM"
windowLayers(152).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Clr 6mm/13mm Arg ! 2665  U=1.32  SC= .48  SHGC=.42  TSOL=.34  TVIS=.68
windowLayers(153).layerName(1) = "LoE SPEC SEL CLEAR 6MM"
windowLayers(153).layerName(2) = "ARGON 13MM"
windowLayers(153).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Tint 6mm/6mm Air ! 2666  U=2.41  SC= .35  SHGC=.31  TSOL=.21  TVIS=.41
windowLayers(154).layerName(1) = "LoE SPEC SEL TINT 6MM"
windowLayers(154).layerName(2) = "AIR 6MM"
windowLayers(154).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Tint 6mm/13mm Air! 2667  U=1.67  SC= .33  SHGC=.29  TSOL=.21  TVIS=.41
windowLayers(155).layerName(1) = "LoE SPEC SEL TINT 6MM"
windowLayers(155).layerName(2) = "AIR 13MM"
windowLayers(155).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Spec Sel Tint 6mm/13mm Arg! 2668  U=1.32  SC= .32  SHGC=.28  TSOL=.21  TVIS=.41
windowLayers(156).layerName(1) = "LoE SPEC SEL TINT 6MM"
windowLayers(156).layerName(2) = "ARGON 13MM"
windowLayers(156).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Abs Bleached 6mm/6mm Air ! 2800  U=2.43  SC= .85  SHGC=.73  TSOL=.64  TVIS=.76
windowLayers(157).layerName(1) = "ECABS-2 BLEACHED 6MM"
windowLayers(157).layerName(2) = "AIR 6MM"
windowLayers(157).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Abs Colored 6mm/6mm Air  ! 2801  U=2.43  SC= .21  SHGC=.18  TSOL=.09  TVIS=.12
windowLayers(158).layerName(1) = "ECABS-2 COLORED 6MM"
windowLayers(158).layerName(2) = "AIR 6MM"
windowLayers(158).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Abs Bleached 6mm/13mm Air! 2802  U=1.78  SC= .86  SHGC=.74  TSOL=.64  TVIS=.76
windowLayers(159).layerName(1) = "ECABS-2 BLEACHED 6MM"
windowLayers(159).layerName(2) = "AIR 13MM"
windowLayers(159).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Abs Colored 6mm/13mm Air ! 2803  U=1.78  SC= .19  SHGC=.20  TSOL=.16  TVIS=.12
windowLayers(160).layerName(1) = "ECABS-2 COLORED 6MM"
windowLayers(160).layerName(2) = "AIR 13MM"
windowLayers(160).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Abs Bleached 6mm/13mm Arg! 2804  U=1.49  SC= .86  SHGC=.74  TSOL=.64  TVIS=.76
windowLayers(161).layerName(1) = "ECABS-2 BLEACHED 6MM"
windowLayers(161).layerName(2) = "ARGON 13MM"
windowLayers(161).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Abs Colored 6mm/13mm Arg ! 2805  U=1.49  SC= .18  SHGC=.15  TSOL=.09  TVIS=.12
windowLayers(162).layerName(1) = "ECABS-2 COLORED 6MM"
windowLayers(162).layerName(2) = "ARGON 13MM"
windowLayers(162).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Ref Bleached 6mm/6mm Air ! 2820  U=2.43  SC= .73  SHGC=.63  TSOL=.55  TVIS=.73
windowLayers(163).layerName(1) = "ECREF-2 BLEACHED 6MM"
windowLayers(163).layerName(2) = "AIR 6MM"
windowLayers(163).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Ref Colored 6mm/6mm Air  ! 2821  U=2.43  SC= .20  SHGC=.17  TSOL=.09  TVIS=.14
windowLayers(164).layerName(1) = "ECREF-2 COLORED 6MM"
windowLayers(164).layerName(2) = "AIR 6MM"
windowLayers(164).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Ref Bleached 6mm/13mm Air! 2822  U=1.78  SC= .74  SHGC=.64  TSOL=.55  TVIS=.73
windowLayers(165).layerName(1) = "ECREF-2 BLEACHED 6MM"
windowLayers(165).layerName(2) = "AIR 13MM"
windowLayers(165).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Ref Colored 6mm/13mm Air ! 2823  U=1.78  SC= .17  SHGC=.15  TSOL=.09  TVIS=.14
windowLayers(166).layerName(1) = "ECREF-2 COLORED 6MM"
windowLayers(166).layerName(2) = "AIR 13MM"
windowLayers(166).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Ref Bleached 6mm/13mm Arg! 2824  U=1.49  SC= .74  SHGC=.64  TSOL=.55  TVIS=.73
windowLayers(167).layerName(1) = "ECREF-2 BLEACHED 6MM"
windowLayers(167).layerName(2) = "ARGON 13MM"
windowLayers(167).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl Elec Ref Colored 6mm/13mm Arg ! 2825  U=1.49  SC= .16  SHGC=.15  TSOL=.09  TVIS=.14
windowLayers(168).layerName(1) = "ECREF-2 COLORED 6MM"
windowLayers(168).layerName(2) = "ARGON 13MM"
windowLayers(168).layerName(3) = "CLEAR 6MM"
'CONSTRUCTION Dbl LoE Elec Abs Bleached 6mm/6mm Air ! 2840  U=2.33  SC= .51  SHGC=.44  TSOL=.34  TVIS=.66
windowLayers(169).layerName(1) = "ECABS-2 BLEACHED 6MM"
windowLayers(169).layerName(2) = "AIR 6MM"
windowLayers(169).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Abs Colored 6mm/6mm Air  ! 2841  U=2.33  SC= .18  SHGC=.16  TSOL=.06  TVIS=.10
windowLayers(170).layerName(1) = "ECABS-2 COLORED 6MM"
windowLayers(170).layerName(2) = "AIR 6MM"
windowLayers(170).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Abs Bleached 6mm/13mm Air! 2842  U=1.64  SC= .59  SHGC=.51  TSOL=.34  TVIS=.66
windowLayers(171).layerName(1) = "ECABS-2 BLEACHED 6MM"
windowLayers(171).layerName(2) = "AIR 13MM"
windowLayers(171).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Abs Colored 6mm/13mm Air ! 2843  U=1.64  SC= .15  SHGC=.13  TSOL=.06  TVIS=.10
windowLayers(172).layerName(1) = "ECABS-2 COLORED 6MM"
windowLayers(172).layerName(2) = "AIR 13MM"
windowLayers(172).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Abs Bleached 6mm/13mm Arg! 2844  U=1.33  SC= .60  SHGC=.52  TSOL=.34  TVIS=.66
windowLayers(173).layerName(1) = "ECABS-2 BLEACHED 6MM"
windowLayers(173).layerName(2) = "ARGON 13MM"
windowLayers(173).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Abs Colored 6mm/13mm Arg ! 2845  U=1.33  SC= .14  SHGC=.12  TSOL=.06  TVIS=.10
windowLayers(174).layerName(1) = "ECABS-2 COLORED 6MM"
windowLayers(174).layerName(2) = "ARGON 13MM"
windowLayers(174).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Ref Bleached 6mm/6mm Air ! 2860  U=2.33  SC= .54  SHGC=.46  TSOL=.32  TVIS=.64
windowLayers(175).layerName(1) = "ECREF-2 BLEACHED 6MM"
windowLayers(175).layerName(2) = "AIR 6MM"
windowLayers(175).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Ref Colored 6mm/6mm Air  ! 2861  U=2.33  SC= .18  SHGC=.16  TSOL=.07  TVIS=.12
windowLayers(176).layerName(1) = "ECREF-2 COLORED 6MM"
windowLayers(176).layerName(2) = "AIR 6MM"
windowLayers(176).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Ref Bleached 6mm/13mm Air! 2862  U=1.64  SC= .55  SHGC=.47  TSOL=.32  TVIS=.64
windowLayers(177).layerName(1) = "ECREF-2 BLEACHED 6MM"
windowLayers(177).layerName(2) = "AIR 13MM"
windowLayers(177).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Ref Colored 6mm/13mm Air ! 2863  U=1.64  SC= .16  SHGC=.14  TSOL=.07  TVIS=.12
windowLayers(178).layerName(1) = "ECREF-2 COLORED 6MM"
windowLayers(178).layerName(2) = "AIR 13MM"
windowLayers(178).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Ref Bleached 6mm/13mm Arg! 2864  U=1.33  SC= .56  SHGC=.48  TSOL=.32  TVIS=.64
windowLayers(179).layerName(1) = "ECREF-2 BLEACHED 6MM"
windowLayers(179).layerName(2) = "ARGON 13MM"
windowLayers(179).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Dbl LoE Elec Ref Colored 6mm/13m Arg  ! 2865  U=1.33  SC= .15  SHGC=.13  TSOL=.07  TVIS=.12
windowLayers(180).layerName(1) = "ECREF-2 COLORED 6MM"
windowLayers(180).layerName(2) = "ARGON 13MM"
windowLayers(180).layerName(3) = "LoE SPEC SEL CLEAR 6MM Rev"
'CONSTRUCTION Trp Clr 3mm/6mm Air   ! 3001  U=2.19  SC= .79  SHGC=.68  TSOL=.60  TVIS=.74
windowLayers(181).layerName(1) = "CLEAR 3MM"
windowLayers(181).layerName(2) = "AIR 6MM"
windowLayers(181).layerName(3) = "CLEAR 3MM"
windowLayers(181).layerName(4) = "AIR 6MM"
windowLayers(181).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp Clr 3mm/13mm Air  ! 3002  U=1.79  SC= .79  SHGC=.68  TSOL=.60  TVIS=.74
windowLayers(182).layerName(1) = "CLEAR 3MM"
windowLayers(182).layerName(2) = "AIR 13MM"
windowLayers(182).layerName(3) = "CLEAR 3MM"
windowLayers(182).layerName(4) = "AIR 13MM"
windowLayers(182).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp Clr 3mm/13mm Arg  ! 3003  U=1.64  SC= .79  SHGC=.68  TSOL=.60  TVIS=.74
windowLayers(183).layerName(1) = "CLEAR 3MM"
windowLayers(183).layerName(2) = "ARGON 13MM"
windowLayers(183).layerName(3) = "CLEAR 3MM"
windowLayers(183).layerName(4) = "ARGON 13MM"
windowLayers(183).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp LoE (e5=.1) Clr 3mm/6mm Air   ! 3601  U=1.81  SC= .67  SHGC=.57  TSOL=.46  TVIS=.70
windowLayers(184).layerName(1) = "CLEAR 3MM"
windowLayers(184).layerName(2) = "AIR 6MM"
windowLayers(184).layerName(3) = "CLEAR 3MM"
windowLayers(184).layerName(4) = "AIR 6MM"
windowLayers(184).layerName(5) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Trp LoE (e5=.1) Clr 3mm/13mm Air  ! 3602  U=1.28  SC= .67  SHGC=.58  TSOL=.46  TVIS=.70
windowLayers(185).layerName(1) = "CLEAR 3MM"
windowLayers(185).layerName(2) = "AIR 13MM"
windowLayers(185).layerName(3) = "CLEAR 3MM"
windowLayers(185).layerName(4) = "AIR 13MM"
windowLayers(185).layerName(5) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Trp LoE (e5=.1) Clr 3mm/13mm Arg  ! 3603  U=1.06  SC= .67  SHGC=.58  TSOL=.46  TVIS=.70
windowLayers(186).layerName(1) = "CLEAR 3MM"
windowLayers(186).layerName(2) = "ARGON 13MM"
windowLayers(186).layerName(3) = "CLEAR 3MM"
windowLayers(186).layerName(4) = "ARGON 13MM"
windowLayers(186).layerName(5) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Trp LoE (e2=e5=.1) Clr 3mm/6mm Air! 3621  U=1.55  SC= .54  SHGC=.47  TSOL=.36  TVIS=.66
windowLayers(187).layerName(1) = "LoE CLEAR 3MM"
windowLayers(187).layerName(2) = "AIR 6MM"
windowLayers(187).layerName(3) = "CLEAR 3MM"
windowLayers(187).layerName(4) = "AIR 6MM"
windowLayers(187).layerName(5) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Trp LoE (e2=e5=.1) Clr 3mm/13mm Air   ! 3622  U=0.99  SC= .55  SHGC=.47  TSOL=.36  TVIS=.66
windowLayers(188).layerName(1) = "LoE CLEAR 3MM"
windowLayers(188).layerName(2) = "AIR 13MM"
windowLayers(188).layerName(3) = "CLEAR 3MM"
windowLayers(188).layerName(4) = "AIR 13MM"
windowLayers(188).layerName(5) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Trp LoE (e2=e5=.1) Clr 3mm/13mm Arg   ! 3623  U=0.77  SC= .55  SHGC=.47  TSOL=.36  TVIS=.66
windowLayers(189).layerName(1) = "LoE CLEAR 3MM"
windowLayers(189).layerName(2) = "ARGON 13MM"
windowLayers(189).layerName(3) = "CLEAR 3MM"
windowLayers(189).layerName(4) = "ARGON 13MM"
windowLayers(189).layerName(5) = "LoE CLEAR 3MM Rev"
'CONSTRUCTION Trp LoE Film (88) Clr 3mm/6mm Air ! 3641  U=1.83  SC= .66  SHGC=.57  TSOL=.48  TVIS=.71
windowLayers(190).layerName(1) = "CLEAR 3MM"
windowLayers(190).layerName(2) = "AIR 6MM"
windowLayers(190).layerName(3) = "COATED POLY-88"
windowLayers(190).layerName(4) = "AIR 6MM"
windowLayers(190).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp LoE Film (88) Clr 3mm/13mm Air! 3642  U=1.32  SC= .67  SHGC=.57  TSOL=.48  TVIS=.71
windowLayers(191).layerName(1) = "CLEAR 3MM"
windowLayers(191).layerName(2) = "AIR 13MM"
windowLayers(191).layerName(3) = "COATED POLY-88"
windowLayers(191).layerName(4) = "AIR 13MM"
windowLayers(191).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp LoE Film (77) Clr 3mm/6mm Air ! 3651  U=1.79  SC= .53  SHGC=.46  TSOL=.38  TVIS=.64
windowLayers(192).layerName(1) = "CLEAR 3MM"
windowLayers(192).layerName(2) = "AIR 6MM"
windowLayers(192).layerName(3) = "COATED POLY-77"
windowLayers(192).layerName(4) = "AIR 6MM"
windowLayers(192).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp LoE Film (77) Clr 3mm/13mm Air! 3652  U=1.26  SC= .54  SHGC=.47  TSOL=.38  TVIS=.64
windowLayers(193).layerName(1) = "CLEAR 3MM"
windowLayers(193).layerName(2) = "AIR 13MM"
windowLayers(193).layerName(3) = "COATED POLY-77"
windowLayers(193).layerName(4) = "AIR 13MM"
windowLayers(193).layerName(5) = "CLEAR 3MM"
'CONSTRUCTION Trp LoE Film (66) Clr 6mm/6mm Air ! 3661  U=1.75  SC= .41  SHGC=.35  TSOL=.26  TVIS=.54
windowLayers(194).layerName(1) = "CLEAR 6MM"
windowLayers(194).layerName(2) = "AIR 6MM"
windowLayers(194).layerName(3) = "COATED POLY-66"
windowLayers(194).layerName(4) = "AIR 6MM"
windowLayers(194).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (66) Clr 6mm/13mm Air! 3662  U=1.23  SC= .42  SHGC=.36  TSOL=.26  TVIS=.54
windowLayers(195).layerName(1) = "CLEAR 6MM"
windowLayers(195).layerName(2) = "AIR 13MM"
windowLayers(195).layerName(3) = "COATED POLY-66"
windowLayers(195).layerName(4) = "AIR 13MM"
windowLayers(195).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (66) Bronze 6mm/6mm Air  ! 3663  U=1.75  SC= .30  SHGC=.26  TSOL=.16  TVIS=.32
windowLayers(196).layerName(1) = "BRONZE 6MM"
windowLayers(196).layerName(2) = "AIR 6MM"
windowLayers(196).layerName(3) = "COATED POLY-66"
windowLayers(196).layerName(4) = "AIR 6MM"
windowLayers(196).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (66) Bronze 6mm/13mm Air ! 3664  U=1.23  SC= .29  SHGC=.25  TSOL=.16  TVIS=.32
windowLayers(197).layerName(1) = "BRONZE 6MM"
windowLayers(197).layerName(2) = "AIR 13MM"
windowLayers(197).layerName(3) = "COATED POLY-66"
windowLayers(197).layerName(4) = "AIR 13MM"
windowLayers(197).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (55) Clr 6mm/6mm Air ! 3671  U=1.74  SC= .35  SHGC=.30  TSOL=.21  TVIS=.45
windowLayers(198).layerName(1) = "CLEAR 6MM"
windowLayers(198).layerName(2) = "AIR 6MM"
windowLayers(198).layerName(3) = "COATED POLY-55"
windowLayers(198).layerName(4) = "AIR 6MM"
windowLayers(198).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (55) Clr 6mm/13m Air ! 3672  U=1.22  SC= .36  SHGC=.31  TSOL=.21  TVIS=.45
windowLayers(199).layerName(1) = "CLEAR 6MM"
windowLayers(199).layerName(2) = "AIR 13MM"
windowLayers(199).layerName(3) = "COATED POLY-55"
windowLayers(199).layerName(4) = "AIR 13MM"
windowLayers(199).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (55) Bronze 6mm/6mm Air  ! 3673  U=1.74  SC= .26  SHGC=.23  TSOL=.13  TVIS=.27
windowLayers(200).layerName(1) = "BRONZE 6MM"
windowLayers(200).layerName(2) = "AIR 6MM"
windowLayers(200).layerName(3) = "COATED POLY-55"
windowLayers(200).layerName(4) = "AIR 6MM"
windowLayers(200).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (55) Bronze 6mm/13mm Air ! 3674  U=1.22  SC= .25  SHGC=.22  TSOL=.13  TVIS=.27
windowLayers(201).layerName(1) = "BRONZE 6MM"
windowLayers(201).layerName(2) = "AIR 13MM"
windowLayers(201).layerName(3) = "COATED POLY-55"
windowLayers(201).layerName(4) = "AIR 13MM"
windowLayers(201).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (44) Bronze 6mm/6mm Air  ! 3681  U=1.74  SC= .23  SHGC=.20  TSOL=.10  TVIS=.22
windowLayers(202).layerName(1) = "BRONZE 6MM"
windowLayers(202).layerName(2) = "AIR 6MM"
windowLayers(202).layerName(3) = "COATED POLY-44"
windowLayers(202).layerName(4) = "AIR 6MM"
windowLayers(202).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (44) Bronze 6mm/13mm Air ! 3682  U=1.21  SC= .22  SHGC=.19  TSOL=.10  TVIS=.22
windowLayers(203).layerName(1) = "BRONZE 6MM"
windowLayers(203).layerName(2) = "AIR 13MM"
windowLayers(203).layerName(3) = "COATED POLY-44"
windowLayers(203).layerName(4) = "AIR 13MM"
windowLayers(203).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (33) Bronze 6mm/6mm Air  ! 3691  U=1.74  SC= .19  SHGC=.16  TSOL=.07  TVIS=.17
windowLayers(204).layerName(1) = "BRONZE 6MM"
windowLayers(204).layerName(2) = "AIR 6MM"
windowLayers(204).layerName(3) = "COATED POLY-33"
windowLayers(204).layerName(4) = "AIR 6MM"
windowLayers(204).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Trp LoE Film (33) Bronze 6mm/13mm Air ! 3692  U=1.20  SC= .17  SHGC=.15  TSOL=.07  TVIS=.17
windowLayers(205).layerName(1) = "BRONZE 6MM"
windowLayers(205).layerName(2) = "AIR 13MM"
windowLayers(205).layerName(3) = "COATED POLY-33"
windowLayers(205).layerName(4) = "AIR 13MM"
windowLayers(205).layerName(5) = "CLEAR 6MM"
'CONSTRUCTION Quadruple LoE Films (88) 3mm/8mm Krypton! 4651  U=0.66  SC= .52  SHGC=.45  TSOL=.34  TVIS=.62
windowLayers(206).layerName(1) = "CLEAR 3MM"
windowLayers(206).layerName(2) = "KRYPTON 8MM"
windowLayers(206).layerName(3) = "COATED POLY-88"
windowLayers(206).layerName(4) = "KRYPTON 3MM"
windowLayers(206).layerName(5) = "COATED POLY-88"
windowLayers(206).layerName(6) = "KRYPTON 8MM"
windowLayers(206).layerName(7) = "CLEAR 3MM"
End Sub

'------------------------------------------------------------------------
' Determine the floors that are active for a new file
'------------------------------------------------------------------------
Sub setActiveFloors()
'floorplan names and count of number of floorplans
numFloorPlans = 1
iFloorPlan(middleFloor).active = True
iFloorPlan(middleFloor).nm = "Main"
If newPlanInfo.isBasement Then
  numFloorPlans = numFloorPlans + 1
  iFloorPlan(basementFloor).active = True
  iFloorPlan(basementFloor).nm = "Basement"
Else
  iFloorPlan(basementFloor).active = False
End If
If newPlanInfo.isFirstDiff Then
  numFloorPlans = numFloorPlans + 1
  iFloorPlan(lowerFloor).active = True
  iFloorPlan(lowerFloor).nm = "Lower"
Else
  iFloorPlan(lowerFloor).active = False
End If
If newPlanInfo.isTopDiff Then
  numFloorPlans = numFloorPlans + 1
  iFloorPlan(topFloor).active = True
  iFloorPlan(topFloor).nm = "Top"
Else
  iFloorPlan(topFloor).active = False
End If
'if more than one floor call the middle floors middle instead of main
If numFloorPlans > 1 Then
  iFloorPlan(middleFloor).nm = "Middle"
End If
End Sub
'------------------------------------------------------------------------
' Create the names for zone, wall and roof objects
'------------------------------------------------------------------------
Sub formZoneWallRoofNames()
Dim i As Integer, j As Integer, k As Integer
Dim nam As String
'name zones
For i = 1 To maxNumFloorPlan
  For j = 1 To numPZone
    nam = iFloorPlan(i).nm & "_Zone_"
    For k = 1 To pZone(j).numZoneCrnrs
      nam = nam & pCorner(pZone(j).crnrs(k)).name
      'if before last item add a underscore to the name
      If k < pZone(j).numZoneCrnrs Then
        nam = nam & "_"
      End If
    Next k
    pZone(j).nm(i) = nam
  Next j
'loop through floorplans for interior and exterior walls
  ' name external walls
  For j = 1 To numPExtWall
    pExtWall(j).nm(i) = iFloorPlan(i).nm & "_ExtWall_" & pCorner(pExtWall(j).startCorner).name & "_" & pCorner(pExtWall(j).endCorner).name
  Next j
  ' name interal walls
  For j = 1 To numPIntWall
    pIntWall(j).nm(i) = iFloorPlan(i).nm & "_IntWall_" & pCorner(pIntWall(j).startCorner).name & "_" & pCorner(pIntWall(j).endCorner).name
  Next j
  ' name windows
  For j = 1 To numPExtWall
    For k = 1 To windowsPerWall
      iWindow(i, j, k).nm = iFloorPlan(i).nm & "_Window_" & pCorner(pExtWall(j).startCorner).name & "_" & pCorner(pExtWall(j).endCorner).name & "_" & Trim(Str(k))
    Next k
  Next j
  ' name doors
  For j = 1 To numPExtWall
    For k = 1 To doorsPerWall
      iDoor(i, j, k).nm = iFloorPlan(i).nm & "_Door_" & pCorner(pExtWall(j).startCorner).name & "_" & pCorner(pExtWall(j).endCorner).name & "_" & Trim(Str(k))
    Next k
  Next j
Next i
'roof name
For j = 1 To numPRoof
  nam = "Roof_"
  For k = 1 To pRoof(j).numRoofCrnrs
    If pRoof(j).crnrs(k) > 0 Then
      nam = nam & pCorner(pRoof(j).crnrs(k)).name
    Else
      nam = nam & pRoofCorner(-pRoof(j).crnrs(k)).name
    End If
    'if before last item add a underscore to the name
    If k < pRoof(j).numRoofCrnrs Then
      nam = nam & "_"
    End If
  Next k
  pRoof(j).nm = nam
Next j
End Sub

'------------------------------------------------------------------------
' Load all expressions and set the order of the user input variables
'------------------------------------------------------------------------
Sub loadExpressions()
Dim isGood As Boolean
Dim Found As Integer
Dim cVarName As String
Dim i As Integer, j As Integer, k As Integer
' normal corners
For i = 1 To numPCorner
  'put expression into evaluation class
  isGood = computeCornerX(i).StoreExpression(pCorner(i).xexpression)
  Debug.Print i, pCorner(i).xexpression, pCorner(i).yexpression
  If Not isGood Then
    MsgBox "Error in template for x expression: " & vbCrLf & _
      pCorner(i).xexpression & vbCrLf & _
      computeCornerX(i).ErrorDescription, vbCritical, "Template Error"
    End
  End If
  ' check input parameters and order of them here
  pCorner(i).xUserInCnt = computeCornerX(i).VarTop
  For j = 1 To computeCornerX(i).VarTop
    Found = 0
    cVarName = computeCornerX(i).VarName(j)
    For k = 1 To numPUserInput
      If pUserInput(k).variable = cVarName Then
        Found = k
        Exit For
      End If
    Next k
    If Found > 0 Then
      pCorner(i).xUserIn(j) = Found
    Else
      MsgBox "User input in expression not found:" & cVarName & " in " & _
      pCorner(i).xexpression, vbCritical, "Parsing error"
    End If
  Next j
  'put expression into evaluation class
  isGood = computeCornerY(i).StoreExpression(pCorner(i).yexpression)
  If Not isGood Then
    MsgBox "Error in template for y expression: " & vbCrLf & _
      pCorner(i).yexpression & vbCrLf & _
      computeCornerY(i).ErrorDescription, vbCritical, "Template Error"
    End
  End If
  ' check input parameters and order of them here
  pCorner(i).yUserInCnt = computeCornerY(i).VarTop
  For j = 1 To computeCornerY(i).VarTop
    Found = 0
    cVarName = computeCornerY(i).VarName(j)
    For k = 1 To numPUserInput
      If pUserInput(k).variable = cVarName Then
        Found = k
        Exit For
      End If
    Next k
    If Found > 0 Then
      pCorner(i).yUserIn(j) = Found
    Else
      MsgBox "User input in expression not found:" & cVarName & " in " & _
      pCorner(i).yexpression, vbCritical, "Parsing error"
    End If
  Next j
Next i
' roof corners
For i = 1 To numPRoofCorner
  'put expression into evaluation class
  isGood = computeRoofCornerX(i).StoreExpression(pRoofCorner(i).xexpression)
  If Not isGood Then
    MsgBox "Error in template for x expression: " & vbCrLf & _
      pRoofCorner(i).xexpression & vbCrLf & _
      computeRoofCornerX(i).ErrorDescription, vbCritical, "Template Error"
    End
  End If
  ' check input parameters and order of them here
  pRoofCorner(i).xUserInCnt = computeRoofCornerX(i).VarTop
  For j = 1 To computeRoofCornerX(i).VarTop
    Found = 0
    cVarName = computeRoofCornerX(i).VarName(j)
    For k = 1 To numPUserInput
      If pUserInput(k).variable = cVarName Then
        Found = k
        Exit For
      End If
    Next k
    If Found > 0 Then
      pRoofCorner(i).xUserIn(j) = Found
    Else
      MsgBox "User input in expression not found:" & cVarName & " in " & _
      pRoofCorner(i).xexpression, vbCritical, "Parsing error"
    End If
  Next j
  'put expression into evaluation class
  isGood = computeRoofCornerY(i).StoreExpression(pRoofCorner(i).yexpression)
  If Not isGood Then
    MsgBox "Error in template for y expression: " & vbCrLf & _
      pRoofCorner(i).yexpression & vbCrLf & _
      computeRoofCornerY(i).ErrorDescription, vbCritical, "Template Error"
    End
  End If
  ' check input parameters and order of them here
  pRoofCorner(i).yUserInCnt = computeRoofCornerY(i).VarTop
  For j = 1 To computeRoofCornerY(i).VarTop
    Found = 0
    cVarName = computeRoofCornerY(i).VarName(j)
    For k = 1 To numPUserInput
      If pUserInput(k).variable = cVarName Then
        Found = k
        Exit For
      End If
    Next k
    If Found > 0 Then
      pRoofCorner(i).yUserIn(j) = Found
    Else
      MsgBox "User input in expression not found:" & cVarName & " in " & _
      pRoofCorner(i).yexpression, vbCritical, "Parsing error"
    End If
  Next j
Next i
' rules
For i = 1 To numPRule
'put expression into evaluation class
isGood = computeRules(i).StoreExpression(pRule(i).Expression)
  If Not isGood Then
    MsgBox "Error in template for rule expression: " & vbCrLf & _
      pRule(i).Expression & vbCrLf & _
      computeRules(i).ErrorDescription, vbCritical, "Template Error"
    End
  End If
  ' check input parameters and order of them here
  pRule(i).UserInCnt = computeRules(i).VarTop
  For j = 1 To computeRules(i).VarTop
    Found = 0
    cVarName = computeRules(i).VarName(j)
    For k = 1 To numPUserInput
      If pUserInput(k).variable = cVarName Then
        Found = k
        Exit For
      End If
    Next k
    If Found > 0 Then
      pRule(i).UserIn(j) = Found
    Else
      MsgBox "User input in expression not found:" & cVarName & " in " & _
      pRule(i).Expression, vbCritical, "Parsing error"
    End If
  Next j
Next i
End Sub

'------------------------------------------------------------------------
' Computes the locations of all points, the length of lines, the areas
' of surfaces.
'------------------------------------------------------------------------
Sub recompute()
Dim result As Single
Dim allRulesOK As Boolean
Dim ruleMsg As String
Dim deltaX As Single, deltaY As Single
Dim sumArea As Single, trapWidth As Single
Dim surfWidth As Single, surfHeight As Single
Dim thisCorner As Integer, nextCorner As Integer
Dim sumWinArea As Single
Dim cosNorthAng As Single, sinNorthAng As Single
Dim i As Integer, j As Integer, k As Integer
On Error Resume Next
' Compute the locations of all corners
For i = 1 To numPCorner
  For j = 1 To pCorner(i).xUserInCnt
    computeCornerX(i).VarValue(j) = pUserInput(pCorner(i).xUserIn(j)).curVal
    Debug.Print "into expression: "; computeCornerX(i).VarName(j), pUserInput(pCorner(i).xUserIn(j)).curVal
  Next j
  pCorner(i).x = computeCornerX(i).Eval
  If Err.Number <> 0 Then
    MsgBox "Evaluation error: " & computeCornerX(i).ErrorDescription, vbCritical, "Template error"
  End If
  For j = 1 To pCorner(i).yUserInCnt
    computeCornerY(i).VarValue(j) = pUserInput(pCorner(i).yUserIn(j)).curVal
  Next j
  pCorner(i).y = computeCornerY(i).Eval
  If Err.Number <> 0 Then
    MsgBox "Evaluation error: " & computeCornerY(i).ErrorDescription, vbCritical, "Template error"
  End If
Next i
For i = 1 To numPRoofCorner
  For j = 1 To pRoofCorner(i).xUserInCnt
    computeRoofCornerX(i).VarValue(j) = pUserInput(pRoofCorner(i).xUserIn(j)).curVal
  Next j
  pRoofCorner(i).x = computeRoofCornerX(i).Eval
  If Err.Number <> 0 Then
    MsgBox "Evaluation error: " & computeRoofCornerX(i).ErrorDescription, vbCritical, "Template error"
  End If
  For j = 1 To pRoofCorner(i).yUserInCnt
    computeRoofCornerY(i).VarValue(j) = pUserInput(pRoofCorner(i).yUserIn(j)).curVal
  Next j
  pRoofCorner(i).y = computeRoofCornerY(i).Eval
  If Err.Number <> 0 Then
    MsgBox "Evaluation error: " & computeRoofCornerY(i).ErrorDescription, vbCritical, "Template error"
  End If
Next i
allRulesOK = True
For i = 1 To numPRule
  For j = 1 To pRule(i).UserInCnt
    computeRules(i).VarValue(j) = pUserInput(pRule(i).UserIn(j)).curVal
  Next j
  result = computeRules(i).Eval
  If result = 1 Then
    pRule(i).isGood = True
  Else
    pRule(i).isGood = False
    allRulesOK = False
  End If
  If Err.Number <> 0 Then
    MsgBox "Evaluation error: " & computeRules(i).ErrorDescription, vbCritical, "Template error"
  End If
Next i
If Not allRulesOK Then
  ruleMsg = ""
  For i = 1 To numPRule
    If Not pRule(i).isGood Then
      ruleMsg = ruleMsg & vbCrLf & vbCrLf & pRule(i).Expression
      For j = 1 To computeRules(i).VarTop
        ruleMsg = ruleMsg & vbCrLf & "  " & computeRules(i).VarName(j) & " = " & Str(computeRules(i).VarValue(j))
      Next j
    End If
  Next i
  MsgBox "One or more validation rules were violated:" & vbCrLf & vbCrLf & ruleMsg, vbCritical, "Invalid Entry"
End If
' compute wall lengths and areas
' length use pythag.
For i = 1 To numPExtWall
  deltaX = pCorner(pExtWall(i).endCorner).x - pCorner(pExtWall(i).startCorner).x
  deltaY = pCorner(pExtWall(i).endCorner).y - pCorner(pExtWall(i).startCorner).y
  pExtWall(i).length = Sqr(deltaY * deltaY + deltaX * deltaX)
  ' use height but remember this could be defaulted
  For j = 1 To maxNumFloorPlan
    If iFloorPlan(j).flr2flr = useNumericDefault Then
      pExtWall(i).area(j) = pExtWall(i).length * iDefault.flr2flr
    Else
      pExtWall(i).area(j) = pExtWall(i).length * iFloorPlan(j).flr2flr
    End If
  Next j
Next i
For i = 1 To numPIntWall
  deltaX = pCorner(pIntWall(i).endCorner).x - pCorner(pIntWall(i).startCorner).x
  deltaY = pCorner(pIntWall(i).endCorner).y - pCorner(pIntWall(i).startCorner).y
  pIntWall(i).length = Sqr(deltaY * deltaY + deltaX * deltaX)
  ' use height but remember this could be defaulted
  For j = 1 To maxNumFloorPlan
    If iFloorPlan(j).flr2flr = useNumericDefault Then
      pIntWall(i).area(j) = pIntWall(i).length * iDefault.flr2flr
    Else
      pIntWall(i).area(j) = pIntWall(i).length * iFloorPlan(j).flr2flr
    End If
  Next j
Next i
'compute zone areas using sum of trapazoids method
For i = 1 To numPZone
  sumArea = 0
  For j = 1 To pZone(i).numZoneCrnrs
    thisCorner = pZone(i).crnrs(j)
    'find the next corner - if on the last corner then the next one is
    'the first corner in the list
    If j = pZone(i).numZoneCrnrs Then
      nextCorner = pZone(i).crnrs(1)
    Else
      nextCorner = pZone(i).crnrs(j + 1)
    End If
    trapWidth = pCorner(nextCorner).x - pCorner(thisCorner).x
    'figure the area of the trapazoid and add it to the area so far
    sumArea = sumArea + trapWidth * (pCorner(nextCorner).y + pCorner(thisCorner).y) / 2
    pZone(i).area = Abs(sumArea) 'could be negative if corners listed in reverse
  Next j
Next i
'compute area of floor by summing zones
sumArea = 0
For i = 1 To numPZone
  sumArea = sumArea + pZone(i).area
Next i
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    iFloorPlan(i).flrArea = sumArea
  End If
Next i
'compute area for building
sumArea = 0
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    sumArea = sumArea + iFloorPlan(i).flrArea * iFloorPlan(i).numFlr
  End If
Next i
iBuilding.floorArea = sumArea
'compute window areas for each exterior wall
' iWindow(maxNumFloorPlan, maxNumPWalls, windowsPerWall) As iWindowType
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    For j = 1 To numPExtWall
      sumWinArea = 0
      For k = 1 To windowsPerWall
        If iWindow(i, j, k).count > 0 Then
          'substitute default if used
          If iWindow(i, j, k).height = useNumericDefault Then
            surfHeight = iDefault.windHeight
          Else
            surfHeight = iWindow(i, j, k).height
          End If
          If iWindow(i, j, k).width = useNumericDefault Then
            surfWidth = iDefault.windWidth
          Else
            surfWidth = iWindow(i, j, k).width
          End If
          sumWinArea = sumWinArea + surfHeight * surfWidth * iWindow(i, j, k).count
        End If
      Next k
      pExtWall(j).glazArea(i) = sumWinArea
      pExtWall(j).perGlaz(i) = (100# * sumWinArea) / pExtWall(j).area(i)
    Next j
  End If
Next i
'calculate the building percent glazing
sumArea = 0
sumWinArea = 0
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    For j = 1 To numPExtWall
      sumArea = sumArea + pExtWall(j).area(i)
      sumWinArea = sumWinArea + pExtWall(j).glazArea(i)
    Next j
  End If
Next i
iBuilding.wallArea = sumArea
iBuilding.glazArea = sumWinArea
iBuilding.perGlaz = 100# * iBuilding.glazArea / iBuilding.wallArea
'calculate the height of the building
surfHeight = 0
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    If iFloorPlan(i).flr2flr = useNumericDefault Then
      surfHeight = surfHeight + iDefault.flr2flr * iFloorPlan(i).numFlr
    Else
      surfHeight = surfHeight + iFloorPlan(i).flr2flr * iFloorPlan(i).numFlr
    End If
  End If
Next i
iBuilding.height = surfHeight
'compute the translated coordinates for corners after building rotation is applied\
cosNorthAng = Cos(iBuilding.northAngle * 3.14159265358979 / 180)
sinNorthAng = Sin(iBuilding.northAngle * 3.14159265358979 / 180)
For i = 1 To numPCorner
  pCorner(i).xTrans = pCorner(i).x * cosNorthAng + pCorner(i).y * sinNorthAng
  pCorner(i).yTrans = -pCorner(i).x * sinNorthAng + pCorner(i).y * cosNorthAng
Next i
For i = 1 To numPRoofCorner
  pRoofCorner(i).xTrans = pRoofCorner(i).x * cosNorthAng + pRoofCorner(i).y * sinNorthAng
  pRoofCorner(i).yTrans = -pRoofCorner(i).x * sinNorthAng + pRoofCorner(i).y * cosNorthAng
Next i
End Sub

'------------------------------------------------------------------------
' This routine is used by the FILE NEW menu item and the NEW button to
' clear all of the arrays used by the program and make the program behave
' like it has just started up.
'------------------------------------------------------------------------
Sub clearAll()
Dim i As Integer, j As Integer, k As Integer
Erase pUserInput
numPUserInput = 0
Erase pRule
numPRule = 0
Erase pCorner
numPCorner = 0
Erase pRoofCorner
numPRoofCorner = 0
Erase pExtWall
numPExtWall = 0
Erase pIntWall
numPIntWall = 0
Erase pZone
numPZone = 0
Erase pRoof
numPRoof = 0
iBuilding.botFloorCons = 0
iBuilding.botFloorInsul = 0
iBuilding.floorArea = 0
iBuilding.floorCons = 0
iBuilding.glazArea = 0
iBuilding.height = 0
iBuilding.perGlaz = 0
iBuilding.planName = ""
iBuilding.roofCons = 0
iBuilding.roofInsul = 0
iBuilding.roofPkHt = 0
iBuilding.duration = 0
iBuilding.wallArea = 0
iBuilding.epVersion = 0
iDefault.doorCons = 0
iDefault.doorHeight = 0
iDefault.doorWidth = 0
iDefault.extWallCons = 0
iDefault.extWallInsul = 0
iDefault.flr2flr = 0
iDefault.style = 0
iDefault.windCons = 0
iDefault.windHeight = 0
iDefault.windOvrhng = 0
iDefault.windSetbck = 0
iDefault.windWidth = 0
Erase iStyle
Erase iFloorPlan
numFloorPlans = 0
Erase iWindow
Erase iDoor
Erase rowData
Erase computeCornerX
Erase computeCornerY
Erase computeRoofCornerX
Erase computeRoofCornerY
Erase computeRules
grdMain.Clear
grdMain.Rows = 1
grdMain.TextMatrix(0, 0) = "Parameter"
grdMain.TextMatrix(0, 1) = "Units"
grdMain.TextMatrix(0, 2) = "Value"
pctMain.Cls
pctMain.Refresh
mainForm.Refresh
End Sub

'------------------------------------------------------------------------
' Save the status of all of the active arrays
'------------------------------------------------------------------------
Sub saveActive()
Dim outFn As Integer
Dim jFloor As Integer
Dim kWall As Integer
Dim lWindow As Integer
Dim mDoor As Integer
Dim i As Integer, j As Integer
Dim c As String
c = ","
outFn = FreeFile
Open curFileNameWithPath For Output As outFn
Print #outFn, "EP-Quick, by Jason Glazer"
Print #outFn, "Version, "; programVersion
'pUserInput
For i = 1 To numPUserInput
  Print #outFn, "pUserInput"; c;
  Print #outFn, pUserInput(i).Description; c;
  Print #outFn, pUserInput(i).variable; c;
  Print #outFn, pUserInput(i).default; c;
  Print #outFn, pUserInput(i).min; c;
  Print #outFn, pUserInput(i).max; c;
  Print #outFn, pUserInput(i).curVal
Next i
'pRule
For i = 1 To numPRule
  Print #outFn, "pRule"; c;
  Print #outFn, pRule(i).Expression
Next i
'newPlanInfo
Print #outFn, "geninfo"; c;
Print #outFn, newPlanInfo.isIPunits
'pCorner
For i = 1 To numPCorner
  Print #outFn, "pCorner"; c;
  Print #outFn, pCorner(i).name; c;
  Print #outFn, pCorner(i).xexpression; c;
  Print #outFn, pCorner(i).yexpression
Next i
'pRoofCorner
For i = 1 To numPRoofCorner
  Print #outFn, "pRoofCorner"; c;
  Print #outFn, pRoofCorner(i).name; c;
  Print #outFn, pRoofCorner(i).xexpression; c;
  Print #outFn, pRoofCorner(i).yexpression
Next i
'pExtWall
For i = 1 To numPExtWall
  Print #outFn, "pExtWall"; c;
  Print #outFn, pExtWall(i).startCorner; c;
  Print #outFn, pExtWall(i).endCorner; c;
  For jFloor = 1 To maxNumFloorPlan
    Print #outFn, pExtWall(i).cons(jFloor); c;
    Print #outFn, pExtWall(i).insul(jFloor); c;
  Next jFloor
  Print #outFn, ""
Next i
'pIntWall
For i = 1 To numPIntWall
  Print #outFn, "pIntWall"; c;
  Print #outFn, pIntWall(i).startCorner; c;
  Print #outFn, pIntWall(i).endCorner
Next i
'pZone
For i = 1 To numPZone
  Print #outFn, "pZone"; c;
  Print #outFn, pZone(i).numZoneCrnrs; c;
  For jFloor = 1 To maxNumFloorPlan
    Print #outFn, pZone(i).style(jFloor); c;
  Next jFloor
  For j = 1 To pZone(i).numZoneCrnrs
    Print #outFn, pZone(i).crnrs(j); c;
  Next j
  Print #outFn, ""
Next i
'pRoof
For i = 1 To numPRoof
  Print #outFn, "pRoof"; c;
  Print #outFn, pRoof(i).numRoofCrnrs; c;
  For j = 1 To pRoof(i).numRoofCrnrs
    Print #outFn, pRoof(i).crnrs(j); c;
  Next j
  Print #outFn, ""
Next i
'iBuilding
Print #outFn, "iBuilding"; c;
Print #outFn, iBuilding.roofCons; c;
Print #outFn, iBuilding.roofInsul; c;
Print #outFn, iBuilding.intWallCons; c;
Print #outFn, iBuilding.floorCons; c;
Print #outFn, iBuilding.botFloorCons; c;
Print #outFn, iBuilding.botFloorInsul; c;
Print #outFn, iBuilding.roofPkHt; c;
Print #outFn, iBuilding.planName; c;
Print #outFn, iBuilding.epVersion; c;
Print #outFn, iBuilding.northAngle; c;
Print #outFn, iBuilding.duration
'iDefault
Print #outFn, "iDefault"; c;
Print #outFn, iDefault.flr2flr; c;
Print #outFn, iDefault.style; c;
Print #outFn, iDefault.extWallCons; c;
Print #outFn, iDefault.extWallInsul; c;
Print #outFn, iDefault.windCons; c;
Print #outFn, iDefault.windWidth; c;
Print #outFn, iDefault.windHeight; c;
Print #outFn, iDefault.windOvrhng; c;
Print #outFn, iDefault.windSetbck; c;
Print #outFn, iDefault.doorCons; c;
Print #outFn, iDefault.doorWidth; c;
Print #outFn, iDefault.doorHeight
'iStyle
For i = 1 To numIStyle
  Print #outFn, "iStyle"; c;
  Print #outFn, iStyle(i).peopDensUse; c;
  Print #outFn, iStyle(i).peopDensNonUse; c;
  Print #outFn, iStyle(i).liteDensUse; c;
  Print #outFn, iStyle(i).liteDensNonUse; c;
  Print #outFn, iStyle(i).eqpDensUse; c;
  Print #outFn, iStyle(i).eqpDensNonUse; c;
  Print #outFn, iStyle(i).weekdayTimeRange; c;
  Print #outFn, iStyle(i).saturdayTimeRange; c;
  Print #outFn, iStyle(i).sundayTimeRange; c;
  Print #outFn, iStyle(i).furnDens
Next i
'iFloorPlan
For i = 1 To maxNumFloorPlan
  If iFloorPlan(i).active Then
    Print #outFn, "iFloorPlan"; c;
    Print #outFn, i; c;
    Print #outFn, iFloorPlan(i).nm; c;
    Print #outFn, iFloorPlan(i).flr2flr; c;
    Print #outFn, iFloorPlan(i).numFlr
  End If
Next i
'iWindow
For jFloor = 1 To maxNumFloorPlan
  For kWall = 1 To numPExtWall
    For lWindow = 1 To windowsPerWall
      If iWindow(jFloor, kWall, lWindow).count > 0 Then
        Print #outFn, "iWindow"; c;
        Print #outFn, jFloor; c;
        Print #outFn, kWall; c;
        Print #outFn, lWindow; c;
        Print #outFn, iWindow(jFloor, kWall, lWindow).cons; c;
        Print #outFn, iWindow(jFloor, kWall, lWindow).width; c;
        Print #outFn, iWindow(jFloor, kWall, lWindow).height; c;
        Print #outFn, iWindow(jFloor, kWall, lWindow).count
'not yet implemented        Print #outFn, iWindow(jFloor, kWall, lWindow).ovrhng; c;
'not yet implemented        Print #outFn, iWindow(jFloor, kWall, lWindow).setbck
      End If
    Next lWindow
  Next kWall
Next jFloor
'iDoor
For jFloor = 1 To maxNumFloorPlan
  For kWall = 1 To numPExtWall
    For mDoor = 1 To doorsPerWall
      If iDoor(jFloor, kWall, mDoor).count > 0 Then
        Print #outFn, "iDoor"; c;
        Print #outFn, jFloor; c;
        Print #outFn, kWall; c;
        Print #outFn, mDoor; c;
        Print #outFn, iDoor(jFloor, kWall, mDoor).cons; c;
        Print #outFn, iDoor(jFloor, kWall, mDoor).width; c;
        Print #outFn, iDoor(jFloor, kWall, mDoor).height; c;
        Print #outFn, iDoor(jFloor, kWall, mDoor).count
      End If
    Next mDoor
  Next kWall
Next jFloor

Close outFn
End Sub

'------------------------------------------------------------------------
' Read the saved active information from the temporary file
'------------------------------------------------------------------------
Sub readActive()
Dim lineFromFile As String
Dim partsOfLine() As String
Dim numOfParts As Integer
Dim numIStyleRead As Integer
Dim jFloor As Integer, kWall As Integer, lWindow As Integer, mDoor As Integer
Dim i As Integer
Dim inFn As Integer
Dim isIncomplete As Boolean
On Error Resume Next
inFn = FreeFile
Open curFileNameWithPath For Input As inFn
'either the end of the file has been reach or the end of the template
'part of the file has been reached
Do While Not EOF(inFn)
  Line Input #inFn, lineFromFile
  'separate the line read into pieces
  partsOfLine = Split(lineFromFile, ",", -1)
  numOfParts = UBound(partsOfLine)
  If numOfParts >= 1 Then
    Select Case partsOfLine(0)
      Case "EP-Quick"
        'skip
      Case "Version"
        'skip
      Case "geninfo"
        If UCase(partsOfLine(1)) = "TRUE" Then
          newPlanInfo.isIPunits = True
        Else
          newPlanInfo.isIPunits = False
        End If
      Case "pUserInput"
        numPUserInput = numPUserInput + 1
        If numPUserInput > maxNumPUserInput Then
          MsgBox "Too many input fields defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pUserInput(numPUserInput).Description = partsOfLine(1)
        pUserInput(numPUserInput).variable = partsOfLine(2)
        pUserInput(numPUserInput).default = CSng(partsOfLine(3))
        pUserInput(numPUserInput).min = CSng(partsOfLine(4))
        pUserInput(numPUserInput).max = CSng(partsOfLine(5))
        pUserInput(numPUserInput).curVal = CSng(partsOfLine(6))
      Case "pRule"
        numPRule = numPRule + 1
        If numPRule > maxNumPRules Then
          MsgBox "Too many rules defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pRule(numPRule).Expression = partsOfLine(1)
        'ignore isGood,userincnt, userin because can be recomputed in loadExpresson
      Case "pCorner"
        numPCorner = numPCorner + 1
        If numPCorner > maxNumPCorners Then
          MsgBox "To many corners defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pCorner(numPCorner).name = partsOfLine(1) 'name of corner
        pCorner(numPCorner).xexpression = partsOfLine(2)
        pCorner(numPCorner).yexpression = partsOfLine(3)
        'ignore x,y,xuserin,yuserin,xuserincnt,yuserincnt because recomputed in LoadExpression
      Case "pRoofCorner"
        numPRoofCorner = numPRoofCorner + 1
        If numPRoofCorner > maxNumPRoofCorners Then
          MsgBox "To many roof corners defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pRoofCorner(numPRoofCorner).name = partsOfLine(1) 'name of roof corner
        pRoofCorner(numPRoofCorner).xexpression = partsOfLine(2)
        pRoofCorner(numPRoofCorner).yexpression = partsOfLine(3)
        'ignore x,y,xuserin,yuserin,xuserincnt,yuserincnt because recomputed in LoadExpression
      Case "pExtWall"
        numPExtWall = numPExtWall + 1
        If numPExtWall > maxNumPWalls Then
          MsgBox "To many extwalls defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pExtWall(numPExtWall).startCorner = CInt(partsOfLine(1))
        pExtWall(numPExtWall).endCorner = CInt(partsOfLine(2))
        If pExtWall(numPExtWall).startCorner < 0 Or pExtWall(numPExtWall).endCorner < 0 Then
          MsgBox "Could not find corner named in extWall", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        For i = 1 To maxNumFloorPlan
          pExtWall(numPExtWall).cons(i) = CInt(partsOfLine(1 + 2 * i))
          pExtWall(numPExtWall).insul(i) = CInt(partsOfLine(2 + 2 * i))
        Next i
        'ignore nm because it is recomputed in the form---Names routine
        'ingore length, area, perglaz,glazarea because determined in the recompute routine
      Case "pIntWall"
        numPIntWall = numPIntWall + 1
        If numPIntWall > maxNumPWalls Then
          MsgBox "To many intwalls defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pIntWall(numPIntWall).startCorner = CInt(partsOfLine(1))
        pIntWall(numPIntWall).endCorner = CInt(partsOfLine(2))
        If pIntWall(numPIntWall).startCorner < 0 Or pIntWall(numPIntWall).endCorner < 0 Then
          MsgBox "Could not find corner named in intWall", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        'ignore nm because it is recomputed in the form---Names routine
        'ingore cons and insul because not used for interior walls
        'ingore length, area, perglaz,glazarea because determined in the recompute routine
      Case "pZone"
        numPZone = numPZone + 1
        If numPZone > maxNumPZones Then
          MsgBox "Too many zones defined in saved", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pZone(numPZone).numZoneCrnrs = CInt(partsOfLine(1))
        For i = 1 To maxNumFloorPlan
          pZone(numPZone).style(i) = partsOfLine(1 + i)
        Next i
        For i = 1 To pZone(numPZone).numZoneCrnrs
          pZone(numPZone).crnrs(i) = CInt(partsOfLine(i + maxNumFloorPlan + 1))
        Next i
        'ignore nm because it is recomputed in the form---Names routine
        'ingore area because determined in the recompute routine
      Case "pRoof"
        numPRoof = numPRoof + 1
        If numPRoof > maxNumPRoof Then
          MsgBox "Too many roof pieces defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        pRoof(numPRoof).numRoofCrnrs = CInt(partsOfLine(1))
        For i = 1 To pRoof(numPRoof).numRoofCrnrs
          pRoof(numPRoof).crnrs(i) = CInt(partsOfLine(i + 1))
        Next i
        'ignore nm because it is recomputed in the form---Names routine
        'ingore area because determined in the recompute routine
      Case "iBuilding"
        'only one instance of iBuilding ever expected
        iBuilding.roofCons = CInt(partsOfLine(1))
        iBuilding.roofInsul = CInt(partsOfLine(2))
        iBuilding.intWallCons = CInt(partsOfLine(3))
        iBuilding.floorCons = CInt(partsOfLine(4))
        iBuilding.botFloorCons = CInt(partsOfLine(5))
        iBuilding.botFloorInsul = CInt(partsOfLine(6))
        iBuilding.roofPkHt = CSng(partsOfLine(7))
        iBuilding.planName = partsOfLine(8)
        If partsOfLine(9) = "1.2.1" Then
          iBuilding.epVersion = epVersion121
        ElseIf partsOfLine(9) = "1.2.2" Then
          iBuilding.epVersion = epVersion122
        ElseIf partsOfLine(9) = "1.2.3" Then
          iBuilding.epVersion = epVersion123
        ElseIf partsOfLine(9) = "1.3.0" Then
          iBuilding.epVersion = epVersion130
        ElseIf partsOfLine(9) = "1.4.0" Then
          iBuilding.epVersion = epVersion140
        ElseIf partsOfLine(9) = "2.0.0" Then
          iBuilding.epVersion = epVersion200
        ElseIf partsOfLine(9) = "2.1.0" Then
          iBuilding.epVersion = epVersion210
        ElseIf partsOfLine(9) = "2.2.0" Then
          iBuilding.epVersion = epVersion220
        ElseIf partsOfLine(9) = "3.0.0" Then
          iBuilding.epVersion = epVersion300
        ElseIf partsOfLine(9) = "3.1.0" Then
          iBuilding.epVersion = epVersion310
        ElseIf partsOfLine(9) = "4.0.0" Then
          iBuilding.epVersion = epVersion400
        Else
          iBuilding.epVersion = CInt(partsOfLine(9))
        End If
        iBuilding.northAngle = CSng(partsOfLine(10))
        iBuilding.duration = CInt(partsOfLine(11))
        'ignore perGlaz,floorarea,wallarea,glazarea,height because determined in recompute routine
      Case "iDefault"
        'only one instance of iDefault ever expected
        iDefault.flr2flr = CSng(partsOfLine(1))
        iDefault.style = CInt(partsOfLine(2))
        iDefault.extWallCons = CInt(partsOfLine(3))
        iDefault.extWallInsul = CInt(partsOfLine(4))
        iDefault.windCons = CInt(partsOfLine(5))
        iDefault.windWidth = CSng(partsOfLine(6))
        iDefault.windHeight = CSng(partsOfLine(7))
        iDefault.windOvrhng = CSng(partsOfLine(8))
        iDefault.windSetbck = CSng(partsOfLine(9))
        iDefault.doorCons = CInt(partsOfLine(10))
        iDefault.doorWidth = CSng(partsOfLine(11))
        iDefault.doorHeight = CSng(partsOfLine(12))
      Case "iStyle"
        numIStyleRead = numIStyleRead + 1
        If numIStyleRead > numIStyle Then
          MsgBox "Too many styles defined in saved file", vbCritical, "Reading Saved File"
          Exit Sub
        End If
        iStyle(numIStyleRead).peopDensUse = CSng(partsOfLine(1))
        iStyle(numIStyleRead).peopDensNonUse = CSng(partsOfLine(2))
        iStyle(numIStyleRead).liteDensUse = CSng(partsOfLine(3))
        iStyle(numIStyleRead).liteDensNonUse = CSng(partsOfLine(4))
        iStyle(numIStyleRead).eqpDensUse = CSng(partsOfLine(5))
        iStyle(numIStyleRead).eqpDensNonUse = CSng(partsOfLine(6))
        iStyle(numIStyleRead).weekdayTimeRange = CSng(partsOfLine(7))
        iStyle(numIStyleRead).saturdayTimeRange = CSng(partsOfLine(8))
        iStyle(numIStyleRead).sundayTimeRange = CSng(partsOfLine(9))
        iStyle(numIStyleRead).furnDens = CSng(partsOfLine(10))
      Case "iFloorPlan"
        jFloor = CInt(partsOfLine(1))
        iFloorPlan(jFloor).nm = partsOfLine(2)
        iFloorPlan(jFloor).active = True
        iFloorPlan(jFloor).flr2flr = CSng(partsOfLine(3))
        iFloorPlan(jFloor).numFlr = CSng(partsOfLine(4))
        numFloorPlans = numFloorPlans + 1
        'ignore floor area since it is computed
      Case "iWindow"
        jFloor = CInt(partsOfLine(1))
        kWall = CInt(partsOfLine(2))
        lWindow = CInt(partsOfLine(3))
        iWindow(jFloor, kWall, lWindow).cons = CInt(partsOfLine(4))
        iWindow(jFloor, kWall, lWindow).width = CSng(partsOfLine(5))
        iWindow(jFloor, kWall, lWindow).height = CSng(partsOfLine(6))
        iWindow(jFloor, kWall, lWindow).count = CSng(partsOfLine(7))
'not yet implementd        iWindow(jFloor, kWall, lWindow).ovrhng = CSng(partsOfLine(8))
'not yet implementd        iWindow(jFloor, kWall, lWindow).setbck = CSng(partsOfLine(9))
        'nm is redetermined
      Case "iDoor"
        jFloor = CInt(partsOfLine(1))
        kWall = CInt(partsOfLine(2))
        mDoor = CInt(partsOfLine(3))
        iDoor(jFloor, kWall, mDoor).cons = CInt(partsOfLine(4))
        iDoor(jFloor, kWall, mDoor).width = CSng(partsOfLine(5))
        iDoor(jFloor, kWall, mDoor).height = CSng(partsOfLine(6))
        iDoor(jFloor, kWall, mDoor).count = CSng(partsOfLine(7))
        'nm is redetermined
      Case Else
        MsgBox "Line of saved file cannot be parsed:" & vbCrLf & vbCrLf & lineFromFile & vbCrLf & partsOfLine(0), vbInformation, "Reading Saved File"
    End Select
  End If
Loop
Close inFn
'check to make sure a minimum number of different objects have been read
isIncomplete = False
If numPUserInput < 1 Then isIncomplete = True
If numPCorner < 3 Then isIncomplete = True
If numPExtWall < 3 Then isIncomplete = True
If numPZone < 1 Then isIncomplete = True
If numIStyleRead < numIStyle Then isIncomplete = True
If numFloorPlans < 1 Then isIncomplete = True
If isIncomplete Then
  MsgBox "File appears incomplete.", vbCritical, "Read File Error"
  Call clearAll
  Call doNewFile
End If
End Sub

'------------------------------------------------------------------------
' Parses a true or false string and returns the boolean value
'------------------------------------------------------------------------
Function stringToTrueFalse(stringIn As String) As Boolean
If UCase(Left(LTrim(stringIn), 1)) = "T" Then
  stringToTrueFalse = True
Else
  stringToTrueFalse = False
End If
End Function

'------------------------------------------------------------------------
' Returns a string that says either true or false depending on boolean
'------------------------------------------------------------------------
Function toTrueFalseString(truthStatement As Boolean) As String
If truthStatement Then
  toTrueFalseString = "TRUE"
Else
  toTrueFalseString = "FALSE"
End If
End Function

'------------------------------------------------------------------------
' open file dialog handler
'------------------------------------------------------------------------
Sub openLocalFileDialog(cancelled As Boolean)
Dim oldFN
oldFN = curFileNameWithPath
cancelled = False
On Error Resume Next
With fsDialog
  If curFilePath = "" Then
    curFilePath = "c:\energyplus\"
    .FileName = "c:\energyplus\untitled.epq"
  Else
    .FileName = curFileName
  End If
  .DialogTitle = "Open EP-Quick File"
  .Filter = "EP-Quick Files (*.epq)|*.epq"
  .FilterIndex = 1
  .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
  .CancelError = True
  .ShowOpen
  If Err.Number <> 0 Then
    curFileNameWithPath = oldFN
    cancelled = True
  Else
    curFileNameWithPath = .FileName
  End If
End With
curFileName = extractFileNameNoExt(curFileNameWithPath)
curFilePath = extractPath(curFileNameWithPath)
End Sub

'------------------------------------------------------------------------
' Display the "Save As" dialog box - used by the
' save and save as routines
'------------------------------------------------------------------------
Sub useFileSaveDialog(cancelled As Boolean)
Dim tempFileName As String
On Error Resume Next
With fsDialog
  If curFilePath = "" Then
    .FileName = "c:\untitled.epq"
  Else
    .FileName = curFileNameWithPath
  End If
  .DialogTitle = "Save EP-Quick File"
  .Filter = "EP-Quick Files (*.epq)|*.epq"
  .FilterIndex = 1
  .Flags = cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  .CancelError = True
  .ShowSave
  If Err.Number <> 0 Then
    curFileNameWithPath = "c:\untitled.epq"
    cancelled = True
  Else
    curFileNameWithPath = .FileName
    cancelled = False
  End If
End With
curFileName = extractFileNameNoExt(curFileNameWithPath)
curFilePath = extractPath(curFileNameWithPath)
End Sub

'------------------------------------------------------------------------
' Removes the path and drive and filename from the
' string and returns only the extension
'------------------------------------------------------------------------
Public Function extractExtension(wholePath As String) As String
Dim periodLoc As Integer
Dim ee As String
periodLoc = InStrRev(wholePath, ".") 'find the last period in path
If periodLoc > 1 Then
  ee = Mid(wholePath, periodLoc + 1)
Else
  ee = ""
End If
extractExtension = UCase(ee)
End Function

'------------------------------------------------------------------------
' Removes the file name from the string and
' returns the drive and path (includes trailing slash)
'------------------------------------------------------------------------
Function extractPath(wp As String) As String
Dim i As Integer, lastSlash As Integer
'scan from the end of the string to find last slash
For i = Len(wp) To 1 Step -1
  If Mid(wp, i, 1) = "\" Then
    lastSlash = i
    Exit For
  End If
Next i
If lastSlash > 0 Then
  extractPath = Left(wp, lastSlash)
Else
  extractPath = ""
End If
End Function

'------------------------------------------------------------------------
' Removes the path and drive and extension from the
' string and returns only the file name
'------------------------------------------------------------------------
Function extractFileNameNoExt(wp As String) As String
Dim i As Integer, lastSlash As Integer, t As String
Dim periodLoc As Integer
'scan from the end of the string to find last slash
For i = Len(wp) To 1 Step -1
  If Mid(wp, i, 1) = "\" Then
    lastSlash = i
    Exit For
  End If
Next i
If lastSlash > 0 Then
  t = Mid(wp, lastSlash + 1)
Else
  t = wp
End If
periodLoc = InStr(t, ".")
If periodLoc > 1 Then
  extractFileNameNoExt = Left(t, periodLoc - 1)
Else
  extractFileNameNoExt = t
End If
End Function

'------------------------------------------------------------------------
' Update the main title bar to have the current file name and status
' of recent save
'------------------------------------------------------------------------
Sub updateWindowTitleBar()
If fileChangedSinceSave Then
  mainForm.Caption = "EP-Quick Freeware - " & curFileNameWithPath & " *"
Else
  mainForm.Caption = "EP-Quick Freeware - " & curFileNameWithPath
End If
End Sub

'------------------------------------------------------------------------
' Windows and doors are saved only when the count is positive.  This
' routine sets the default values for all windows and doors prior to
' being opened to default so that they display the Use Default when
' opening a file.
'------------------------------------------------------------------------
Sub setWindowsDoorsToDefault()
Dim jFloor As Integer
Dim kWall As Integer
Dim lWindow As Integer
Dim mDoor As Integer
'iWindow
For jFloor = 1 To maxNumFloorPlan
  For kWall = 1 To maxNumPWalls
    For lWindow = 1 To windowsPerWall
      iWindow(jFloor, kWall, lWindow).width = useNumericDefault
      iWindow(jFloor, kWall, lWindow).height = useNumericDefault
      iWindow(jFloor, kWall, lWindow).ovrhng = useNumericDefault
      iWindow(jFloor, kWall, lWindow).setbck = useNumericDefault
    Next lWindow
  Next kWall
Next jFloor
'iDoor
For jFloor = 1 To maxNumFloorPlan
  For kWall = 1 To maxNumPWalls
    For mDoor = 1 To doorsPerWall
      iDoor(jFloor, kWall, mDoor).width = useNumericDefault
      iDoor(jFloor, kWall, mDoor).height = useNumericDefault
    Next mDoor
  Next kWall
Next jFloor
End Sub

'------------------------------------------------------------------------
' Create the IDF File
'------------------------------------------------------------------------
Sub doCreateIDF()
Dim s As String
Dim idfLocMsg As String
If fileChangedSinceSave Then
  MsgBox "Please save your file before making an EnergyPlus IDF file", vbInformation, "Make IDF Error"
  Exit Sub
End If
If fileIsUntitled Then
  MsgBox "Please save your file with a specific file name before making an EnergyPlus IDF file", vbInformation, "Make IDF Error"
  Exit Sub
End If
If warnUserWhereisIDF Then
  idfLocMsg = "The IDF file will be created in the same directory as the EPQ file" _
  & " with the same name as the EPQ file but with the extension IDF." _
  & vbCrLf & vbCrLf & "The file name is:" & vbCrLf & vbCrLf _
  & "  " & curFilePath & curFileName & ".IDF" & vbCrLf & vbCrLf _
  & "Run this file in EnergyPlus by using EP-Launch included with EnergyPlus. You must " _
  & "download and install EnergyPlus from www.energyplus.gov to run a simulation" & vbCrLf & vbCrLf _
  & "This message is shown only once."
  MsgBox idfLocMsg, vbInformation, "IDF File Location"
  warnUserWhereisIDF = False
End If
'first find out which zones are associated with each wall
Call associateWallsWithZones
'calculate the heights of the floors
Call determineFloorHeights
'next convert distances and other values from IP to SI units (i.e., feet -> meters)
Call convertToSI
'find a new file handle
idfFileHandle = FreeFile
Open curFilePath & curFileName & ".IDF" For Output As idfFileHandle
Call createIDFheaders
Call createIDFconstructions
Call createIDFwindowLayers
Call createIDFvertSurfaces
Call createIDFhorizSurfaces
Call createIDFspacegains
Call createIDFschedules
Close idfFileHandle
End Sub

'------------------------------------------------------------------------
' Determine the height of each floor in the building. Since a building
' could have floors that are repeated. The floorheights are based on
' the average height that floor should be located.
'------------------------------------------------------------------------
Sub determineFloorHeights()
Dim iFloor As Integer
Dim curFloorHeight As Single
Dim curFlr2Flr As Single
curFloorHeight = 0
For iFloor = 1 To maxNumFloorPlan
  If iFloorPlan(iFloor).active Then
    iFloorPlan(iFloor).heightOfFloor = curFloorHeight
    If iFloorPlan(iFloor).flr2flr = useNumericDefault Then
      curFlr2Flr = iDefault.flr2flr
    Else
      curFlr2Flr = iFloorPlan(iFloor).flr2flr
    End If
    curFloorHeight = curFloorHeight + curFlr2Flr 'no longer use   * iFloorPlan(iFloor).numFlr
  End If
Next iFloor
End Sub

'------------------------------------------------------------------------
' Creates the top of the IDF file including the following objects:
'    version
'    building
'    timestep in hour
'    surfacegeometry
'------------------------------------------------------------------------
Sub createIDFheaders()
Print #idfFileHandle, "! File Created by EP-Quick version: "; programVersion; "  at: "; Date; "  "; Time
'version
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Version"
Else
  objIDF "VERSION"
End If
strIDF "Version Identifier", listOfChoices(iBuilding.epVersion), True
'building
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Building"
Else
  objIDF "BUILDING"
End If
strIDF "Building Name", "None"
numIDF "North Axis {deg}", 0
strIDF "Terrain", "Suburbs"
numIDF "Loads Convergence Tolerance Value {W}", 0.04
numIDF "Temperature Convergence Tolerance Value {deltaC}", 0.4
'strIDF "Solar Distribution", "FullInteriorAndExterior" - changed for NOE bug 8/8/2004
'strIDF "Solar Distribution", "MinimalShadowing" - changed because EnergyPlus team considres Minimal not adequate
strIDF "Solar Distribution", "FullInteriorAndExterior"
numIDF "Maximum Number of Warmup Days", 20, True
'TIMESTEP IN HOUR
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Timestep"
Else
  objIDF "TIMESTEP IN HOUR"
End If
numIDF "Time Step in Hour", 4, True
'SURFACEGEOMETRY
If iBuilding.epVersion >= epVersion300 Then
  objIDF "GlobalGeometryRules"
Else
  objIDF "SURFACEGEOMETRY"
End If
strIDF "Surface Starting Position", "UpperLeftCorner"
strIDF "Vertex Entry", "CounterClockWise"
If iBuilding.epVersion >= epVersion310 Then
  strIDF "Coordinate System", "World"
  strIDF "Daylighting Reference Point Coordinate System", "World"
  strIDF "Rectangular Surface Coordinate System", "World", True
Else
  strIDF "Coordinate System", "WorldCoordinateSystem", True
End If
'annual simulation
If iBuilding.duration = durationAnnual Then
  'RUN PERIOD
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "RunPeriod"
  Else
    objIDF "RUNPERIOD"
  End If
  If iBuilding.epVersion >= epVersion400 Then
    strIDF "Name", "FullAnnualRun"
  End If
  numIDF "Begin Month", 1
  numIDF "Begin Day Of Month", 1
  numIDF "End Month", 12  'commented out only for debugging
  numIDF "End Day Of Month", 31  'commented out only for debugging
  strIDF "Day Of Week For Start Day - Use weather file", ""
  strIDF "Use WeatherFile Holidays/Special Days", "Yes"
  strIDF "Use WeatherFile DaylightSavingPeriod", "Yes"
  strIDF "Apply Weekend Holiday Rule", "Yes"
  strIDF "Use WeatherFile Rain Indicators", "Yes"
  strIDF "Use WeatherFile Snow Indicators", "Yes", True
Else ' design day simulation
  'must have location if not using weather file
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Site:Location"
  Else
    objIDF "LOCATION"
  End If
  strIDF "Name", "Midwest City"
  numIDF "Latitude", 42
  numIDF "Longitude", -88
  numIDF "Time Zone", -6
  numIDF "Elevation", 200, True
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "SizingPeriod:DesignDay"
  Else
    objIDF "DESIGNDAY"
  End If
  strIDF "Name", "Cold Midwest Winter"
  numIDF "Maximum Dry Bulb Temperature", -18
  numIDF "Daily Temperature Range", 0
  numIDF "Humidity Indicating Temp at Max", -18
  numIDF "Barometric Pressure", 99872
  numIDF "Wind Speed", 5.5
  numIDF "Wind Direction", 325
  numIDF "Sky Clearness", 0
  numIDF "Rain Indicator", 0
  numIDF "Snow Indicator", 0
  numIDF "Day of Month", 21
  numIDF "Month", 1
  strIDF "Day type", "Monday"
  numIDF "Daylight Savings Time Indicator", 0
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Humidity Indicating Temp Type", "WetBulb", True
  Else
    strIDF "Humidity Indicating Temp Type", "Wet-Bulb", True
  End If
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "SizingPeriod:DesignDay"
  Else
    objIDF "DESIGNDAY"
  End If
  strIDF "Name", "Hot Midwest Summer"
  numIDF "Maximum Dry Bulb Temperature", 33
  numIDF "Daily Temperature Range", 13
  numIDF "Humidity Indicating Temp at Max", 23
  numIDF "Barometric Pressure", 99433
  numIDF "Wind Speed", 3.8
  numIDF "Wind Direction", 210
  numIDF "Sky Clearness", 0.98
  numIDF "Rain Indicator", 0
  numIDF "Snow Indicator", 0
  numIDF "Day of Month", 21
  numIDF "Month", 7
  strIDF "Day type", "Monday"
  numIDF "Daylight Savings Time Indicator", 0
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Humidity Indicating Temp Type", "WetBulb", True
  Else
    strIDF "Humidity Indicating Temp Type", "Wet-Bulb", True
  End If
End If
'INSIDE CONVECTION ALGORITHM
If iBuilding.epVersion >= epVersion300 Then
  objIDF "SurfaceConvectionAlgorithm:Inside"
Else
  objIDF "INSIDE CONVECTION ALGORITHM"
End If
strIDF "Type", "Detailed", True
'OUTSIDE CONVECTION ALGORITHM
If iBuilding.epVersion >= epVersion300 Then
  objIDF "SurfaceConvectionAlgorithm:Outside"
Else
  objIDF "OUTSIDE CONVECTION ALGORITHM"
End If
strIDF "Type", "Detailed", True
'SOLUTION ALGORITHM
If iBuilding.epVersion >= epVersion300 Then
  objIDF "HeatBalanceAlgorithm"
Else
  objIDF "SOLUTION ALGORITHM"
End If
strIDF "Type", "CTF", True
'GROUNDTEMPERATURES
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Site:GroundTemperature:BuildingSurface"
Else
  objIDF "GROUNDTEMPERATURES"
End If
numIDF "Jan", 18
numIDF "Feb", 18
numIDF "Mar", 18
numIDF "Apr", 18
numIDF "Map", 18
numIDF "Jun", 18
numIDF "Jul", 18
numIDF "Aug", 18
numIDF "Sep", 18
numIDF "Oct", 18
numIDF "Nov", 18
numIDF "Dec", 18, True
'  ScheduleType,
If iBuilding.epVersion >= epVersion300 Then
  objIDF "ScheduleTypeLimits"
Else
  objIDF "SCHEDULETYPE"
End If
strIDF "Name", "Fraction"
If iBuilding.epVersion >= epVersion400 Then
  strIDF "Lower Limit Value", "0.0"
  strIDF "Upper Limit Value", "1.0"
Else
  strIDF "Range", " 0.0 : 1.0"
End If
strIDF "Type", "CONTINUOUS", True
If iBuilding.epVersion >= epVersion300 Then
  objIDF "ScheduleTypeLimits"
Else
  objIDF "SCHEDULETYPE"
End If
strIDF "Name", "AnyNumber", True
'Standard reports
If iBuilding.epVersion >= epVersion400 Then
  objIDF "Output:Surfaces:List"
  strIDF "Report Type", "DetailsWithVertices", True
  objIDF "Output:Surfaces:Drawing"
  strIDF "Report Type", "DXF", True
  objIDF "Output:VariableDictionary"
  strIDF "Key Field", "regular", True
  objIDF "OutputControl:Table:Style"
  strIDF "type", "HTML", True
Else 'old style where each report was called from Output:Reports
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Output:Reports"
  Else
    objIDF "REPORT"
  End If
  strIDF "type", "surfaces"
  strIDF "subtype", "detailswithvertices", True
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Output:Reports"
  Else
    objIDF "REPORT"
  End If
  strIDF "type", "surfaces"
  strIDF "subtype", "DXF", True
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Output:Reports"
    strIDF "type", "VariableDictionary", True
    objIDF "Output:Reports"
    strIDF "type", "construction", True
    objIDF "OutputControl:Table:Style"
    strIDF "type", "HTML", True
  Else
    objIDF "REPORT"
    strIDF "type", "variable dictionary", True
    objIDF "REPORT"
    strIDF "type", "construction", True
    objIDF "REPORT:TABLE:STYLE"
    strIDF "type", "HTML", True
  End If
End If
If iBuilding.duration = durationAnnual Then
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Output:Table:SummaryReports"
  Else
    objIDF "REPORT:TABLE:PREDEFINED"
  End If
  If iBuilding.epVersion >= epVersion310 Then
    strIDF "Report 1 Name", "AllSummary", True
  ElseIf iBuilding.epVersion >= epVersion300 Then
    strIDF "Report 1 Name", "AnnualBuildingUtilityPerformanceSummary", True
  Else
    strIDF "Type", "Annual Building Utility Performance Summary", True
  End If
End If
'Monthly report on building cooling loads
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Output:Table:Monthly"
Else
  objIDF "REPORT:TABLE:MONTHLY"
End If
strIDF "Title", "Building Loads - Cooling"
numIDF "DigitsAfterDecimal", 2
strIDF "Variable", "Zone/Sys Sensible Cooling Energy"
strIDF "Aggregation", "SumOrAverage"
strIDF "Variable", "Zone/Sys Sensible Cooling Rate"
strIDF "Aggregation", "Maximum"
strIDF "Variable", "Outdoor Dry Bulb"
If iBuilding.epVersion >= epVersion300 Then
  strIDF "Aggregation", "ValueWhenMaximumOrMinimum"
Else
  strIDF "Aggregation", "ValueWhenMaxMin"
End If
strIDF "Variable", "Outdoor Wet Bulb"
If iBuilding.epVersion >= epVersion300 Then
  strIDF "Aggregation", "ValueWhenMaximumOrMinimum"
Else
  strIDF "Aggregation", "ValueWhenMaxMin"
End If
strIDF "Variable", "Zone Total Internal Latent Gain"
strIDF "Aggregation", "SumOrAverage"
strIDF "Variable", "Zone Total Internal Latent Gain"
strIDF "Aggregation", "Maximum"
strIDF "Variable", "Outdoor Dry Bulb"
If iBuilding.epVersion >= epVersion300 Then
  strIDF "Aggregation", "ValueWhenMaximumOrMinimum"
Else
  strIDF "Aggregation", "ValueWhenMaxMin"
End If
strIDF "Variable", "Outdoor Wet Bulb"
If iBuilding.epVersion >= epVersion300 Then
  strIDF "Aggregation", "ValueWhenMaximumOrMinimum", True
Else
  strIDF "Aggregation", "ValueWhenMaxMin", True
End If
'Monthly report on building heating loads
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Output:Table:Monthly"
Else
  objIDF "REPORT:TABLE:MONTHLY"
End If
strIDF "Title", "Building Loads - Heating"
numIDF "DigitsAfterDecimal", 2
strIDF "Variable", "Zone/Sys Sensible Heating Energy"
strIDF "Aggregation", "SumOrAverage"
strIDF "Variable", "Zone/Sys Sensible Heating Rate"
strIDF "Aggregation", "Maximum"
strIDF "Variable", "Outdoor Dry Bulb"
If iBuilding.epVersion >= epVersion300 Then
  strIDF "Aggregation", "ValueWhenMaximumOrMinimum", True
Else
  strIDF "Aggregation", "ValueWhenMaxMin", True
End If
End Sub

'------------------------------------------------------------------------
' Creates the portion of the IDF file that is related to surfaces and
' subsurfaces (windows and doors) including objects:
'    surface:heattransfer
'    surface:heattransfer:sub (in called subroutine)
'------------------------------------------------------------------------
Sub createIDFvertSurfaces()
Dim curFloorToCeil As Single
Dim iWall As Integer
Dim jFloor As Integer
'surface:heattransfter
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    If iFloorPlan(jFloor).flr2flr = useNumericDefault Then
      curFloorToCeil = iDefault.flr2flrSI
    Else
      curFloorToCeil = iFloorPlan(jFloor).flr2flrSI
    End If
    'exterior walls
    For iWall = 1 To numPExtWall
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "BuildingSurface:Detailed"
      Else
        objIDF "SURFACE:HEATTRANSFER"
      End If
      strIDF "Name", pExtWall(iWall).nm(jFloor)
      strIDF "Surface Type", "Wall"
      strIDF "Construction Name", constInsulCombo(pExtWall(iWall).consInsul(jFloor)).nm
      strIDF "Inside Face Environment", pZone(pExtWall(iWall).zone1).nm(jFloor)
      If iBuilding.epVersion >= epVersion300 Then
        strIDF "Outside Face Environment", "Outdoors"
      Else
        strIDF "Outside Face Environment", "ExteriorEnvironment"
      End If
      strIDF "Outside Face Environment Object", ""
      strIDF "Sun Exposure", "SunExposed"
      strIDF "Wind Exposure", "WindExposed"
      numIDF "View Factor to Ground", 0.5
      numIDF "Number of Surface Vertex Groups", 4
      numIDF "Vertex 1 X-coordinate", pCorner(pExtWall(iWall).endCorner).xTransSI
      numIDF "Vertex 1 Y-coordinate", pCorner(pExtWall(iWall).endCorner).yTransSI
      numIDF "Vertex 1 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI + curFloorToCeil
      numIDF "Vertex 2 X-coordinate", pCorner(pExtWall(iWall).endCorner).xTransSI
      numIDF "Vertex 2 Y-coordinate", pCorner(pExtWall(iWall).endCorner).yTransSI
      numIDF "Vertex 2 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI
      numIDF "Vertex 3 X-coordinate", pCorner(pExtWall(iWall).startCorner).xTransSI
      numIDF "Vertex 3 Y-coordinate", pCorner(pExtWall(iWall).startCorner).yTransSI
      numIDF "Vertex 3 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI
      numIDF "Vertex 4 X-coordinate", pCorner(pExtWall(iWall).startCorner).xTransSI
      numIDF "Vertex 4 Y-coordinate", pCorner(pExtWall(iWall).startCorner).yTransSI
      numIDF "Vertex 4 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI + curFloorToCeil, True
      Call createIDFWindowsDoors(jFloor, iWall)
    Next iWall
    'Interior walls (normal and mirrored copy)
    For iWall = 1 To numPIntWall
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "BuildingSurface:Detailed"
      Else
        objIDF "SURFACE:HEATTRANSFER"
      End If
      strIDF "Name", pIntWall(iWall).nm(jFloor)
      strIDF "Surface Type", "Wall"
      strIDF "Construction Name", "IntWallConstruction"
      strIDF "Inside Face Environment", pZone(pIntWall(iWall).zone1).nm(jFloor)
      If iBuilding.epVersion >= epVersion300 Then
        strIDF "Outside Face Environment", "Surface"
      Else
        strIDF "Outside Face Environment", "OtherZoneSurface"
      End If
      strIDF "Outside Face Environment Object", pIntWall(iWall).nm(jFloor) & "_mirror"
      strIDF "Sun Exposure", "NoSun"
      strIDF "Wind Exposure", "NoWind"
      numIDF "View Factor to Ground", 0.5
      numIDF "Number of Surface Vertex Groups", 4
      numIDF "Vertex 1 X-coordinate", pCorner(pIntWall(iWall).endCorner).xTransSI
      numIDF "Vertex 1 Y-coordinate", pCorner(pIntWall(iWall).endCorner).yTransSI
      numIDF "Vertex 1 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI + curFloorToCeil
      numIDF "Vertex 2 X-coordinate", pCorner(pIntWall(iWall).endCorner).xTransSI
      numIDF "Vertex 2 Y-coordinate", pCorner(pIntWall(iWall).endCorner).yTransSI
      numIDF "Vertex 2 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI
      numIDF "Vertex 3 X-coordinate", pCorner(pIntWall(iWall).startCorner).xTransSI
      numIDF "Vertex 3 Y-coordinate", pCorner(pIntWall(iWall).startCorner).yTransSI
      numIDF "Vertex 3 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI
      numIDF "Vertex 4 X-coordinate", pCorner(pIntWall(iWall).startCorner).xTransSI
      numIDF "Vertex 4 Y-coordinate", pCorner(pIntWall(iWall).startCorner).yTransSI
      numIDF "Vertex 4 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI + curFloorToCeil, True
      'repeated interior wall with opposite zones specified and corners reversed (mirror)
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "BuildingSurface:Detailed"
      Else
        objIDF "SURFACE:HEATTRANSFER"
      End If
      strIDF "Name", pIntWall(iWall).nm(jFloor) & "_mirror"
      strIDF "Surface Type", "Wall"
      strIDF "Construction Name", "IntWallConstruction"
      strIDF "Inside Face Environment", pZone(pIntWall(iWall).zone2).nm(jFloor)
      If iBuilding.epVersion >= epVersion300 Then
        strIDF "Outside Face Environment", "Surface"
      Else
        strIDF "Outside Face Environment", "OtherZoneSurface"
      End If
      strIDF "Outside Face Environment Object", pIntWall(iWall).nm(jFloor)
      strIDF "Sun Exposure", "NoSun"
      strIDF "Wind Exposure", "NoWind"
      numIDF "View Factor to Ground", 0.5
      numIDF "Number of Surface Vertex Groups", 4
      numIDF "Vertex 1 X-coordinate", pCorner(pIntWall(iWall).startCorner).xTransSI
      numIDF "Vertex 1 Y-coordinate", pCorner(pIntWall(iWall).startCorner).yTransSI
      numIDF "Vertex 1 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI + curFloorToCeil
      numIDF "Vertex 2 X-coordinate", pCorner(pIntWall(iWall).startCorner).xTransSI
      numIDF "Vertex 2 Y-coordinate", pCorner(pIntWall(iWall).startCorner).yTransSI
      numIDF "Vertex 2 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI
      numIDF "Vertex 3 X-coordinate", pCorner(pIntWall(iWall).endCorner).xTransSI
      numIDF "Vertex 3 Y-coordinate", pCorner(pIntWall(iWall).endCorner).yTransSI
      numIDF "Vertex 3 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI
      numIDF "Vertex 4 X-coordinate", pCorner(pIntWall(iWall).endCorner).xTransSI
      numIDF "Vertex 4 Y-coordinate", pCorner(pIntWall(iWall).endCorner).yTransSI
      numIDF "Vertex 4 Z-coordinate", iFloorPlan(jFloor).heightOfFloorSI + curFloorToCeil, True
    Next iWall
  End If
Next jFloor
End Sub


'------------------------------------------------------------------------
' Called by createSurfaces for adding the windows and the doors
' Creates:   surface:heattransfer:sub
'------------------------------------------------------------------------
Sub createIDFWindowsDoors(floorIndex As Integer, wallIndex As Integer)
Dim kSub As Integer, mCnt As Integer 'subsurfaces - windows and doors
Dim curFloorToCeil As Single
Dim cntSubsurf As Integer   'count of subsurfaces on the current wall
Dim widthSubsurf As Single  'total width of all subsurfaces on the current wall
Dim curSubWidth As Single   'value of the subsurface width - default values are included
Dim curSubHeight As Single  'value of the subsurface height - default values are accounted for
Dim remainingWallWidth As Single  'width of the wall that has no window or door
Dim gapBetweenSubsurf As Single   'the gap between all of the subsurfaces on the wall
Dim xSlope As Single    'used in determining location of the windows and doors
Dim ySlope As Single    'used in determining location of the windows and doors
Dim L12 As Single       'used in determining location of the windows and doors
Dim curL As Single      'position along the wall for window and door placement
Dim subRightX As Single, subLeftX As Single
Dim subRightY As Single, subLeftY As Single
Dim x1 As Single, y1 As Single 'used in determining location of the windows and doors
Dim x2 As Single, y2 As Single 'used in determining location of the windows and doors
Dim sillHeight As Single       'bottom of window sill or bottom of door
Dim curFloorHeight As Single
Dim curWinCons As Integer
'determine the floor to ceiling height
If iFloorPlan(floorIndex).flr2flr = useNumericDefault Then
  curFloorToCeil = iDefault.flr2flrSI
Else
  curFloorToCeil = iFloorPlan(floorIndex).flr2flrSI
End If
curFloorHeight = iFloorPlan(floorIndex).heightOfFloorSI
'prior to doing windows and doors add up the number of different windows and doors
'and add up the widths of them so that the gaps between them can be computed.  For
'now producing multiple windows using multipliers so they will not display in DXF
'file perfectly.  Could also create every single instance of a window and door but
'that would add to complexity of the file that is probably not justified.
'
'To fix the problem of using solar distribtion with full interior and exterior
'setting the window multiplier cannot be used.  Because of this a window is now
'going to be defined for each one
cntSubsurf = 0
widthSubsurf = 0
' windows - count and accumulate width
For kSub = 1 To windowsPerWall
  If iWindow(floorIndex, wallIndex, kSub).count > 0 Then
    curSubWidth = iWindow(floorIndex, wallIndex, kSub).widthSI
    If iWindow(floorIndex, wallIndex, kSub).width = useNumericDefault Then curSubWidth = iDefault.windWidthSI
    widthSubsurf = widthSubsurf + curSubWidth * iWindow(floorIndex, wallIndex, kSub).count
    'get rid of using multiplier so can use fullexterior solar distribution
    cntSubsurf = cntSubsurf + iWindow(floorIndex, wallIndex, kSub).count
  End If
Next kSub
' doors - count and accumulate width
For kSub = 1 To doorsPerWall
  If iDoor(floorIndex, wallIndex, kSub).count > 0 Then
    curSubWidth = iDoor(floorIndex, wallIndex, kSub).widthSI
    If iDoor(floorIndex, wallIndex, kSub).width = useNumericDefault Then curSubWidth = iDefault.doorWidthSI
    widthSubsurf = widthSubsurf + curSubWidth * iWindow(floorIndex, wallIndex, kSub).count
    'use the count because no door multiplier is available
    cntSubsurf = cntSubsurf + iDoor(floorIndex, wallIndex, kSub).count
  End If
Next kSub
'now that we know how many subsurfaces their are and their total width
'we will calculate how much they should be spread out on the wall
If cntSubsurf > 0 Then
  remainingWallWidth = pExtWall(wallIndex).lengthSI - widthSubsurf
  gapBetweenSubsurf = remainingWallWidth / (cntSubsurf + 1)
  'to find the location of the point in 3D space based on the distance along the wall
  'we need to do some geometry calcs.
  '
  'say x1 y1 are one corner, x2 y2 are the second corner and L is the is the length
  'along the wall that we know we are trying to find xL and yL from those known variables.
  '
  'assuming the length of the wall is L12 = sqrt((x2-x1)^2 + (y2-y1)^2)
  '
  'we know that the ratios L/L12 = (xL - x1)/(x2 - x1) = (yL - y1)/(y2 - y1)
  '
  'rearranging we get
  '
  ' xL = (L/L12)*(x2 - x1) + x1
  ' yL = (L/L12)*(y2 - y1) + y1
  '
  'further, grouping the constants together as much as possible
  '
  ' xL = L * ((x2 - x1)/L12) + x1
  ' yL = L * ((y2 - y1)/L12) + y1
  '
  ' we can call ((x2 - x1)/L12) the xSlope and ((y2 - y1)/L12) the ySlope
  '
  'Their might be easier ways to do this but this seems simple enough
  x1 = pCorner(pExtWall(wallIndex).startCorner).xTransSI
  y1 = pCorner(pExtWall(wallIndex).startCorner).yTransSI
  x2 = pCorner(pExtWall(wallIndex).endCorner).xTransSI
  y2 = pCorner(pExtWall(wallIndex).endCorner).yTransSI
  L12 = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
  xSlope = (x2 - x1) / L12
  ySlope = (y2 - y1) / L12
  'set the position on the wall "L" to one gap
  curL = gapBetweenSubsurf
  'define windows
  For kSub = 1 To windowsPerWall
    If iWindow(floorIndex, wallIndex, kSub).count > 0 Then
      curSubWidth = iWindow(floorIndex, wallIndex, kSub).widthSI
      If iWindow(floorIndex, wallIndex, kSub).width = useNumericDefault Then curSubWidth = iDefault.windWidthSI
      curSubHeight = iWindow(floorIndex, wallIndex, kSub).heightSI
      If iWindow(floorIndex, wallIndex, kSub).height = useNumericDefault Then curSubHeight = iDefault.windHeightSI
      For mCnt = 1 To iWindow(floorIndex, wallIndex, kSub).count
        'window is always positioned at the center line of the wall (relative to the floor)
        sillHeight = (curFloorToCeil - curSubHeight) / 2
        'define the right edge of the window position than the left edge
        subRightX = curL * xSlope + x1
        subRightY = curL * ySlope + y1
        curL = curL + curSubWidth 'add the width of the window
        subLeftX = curL * xSlope + x1
        subLeftY = curL * ySlope + y1
        curL = curL + gapBetweenSubsurf 'add the gap bewteen windows
        If iBuilding.epVersion >= epVersion300 Then
          objIDF "FenestrationSurface:Detailed"
        Else
          objIDF "SURFACE:HEATTRANSFER:SUB"
        End If
        strIDF "Name", iWindow(floorIndex, wallIndex, kSub).nm & "_" & Trim(Str(mCnt))
        strIDF "Type", "Window"
        curWinCons = iWindow(floorIndex, wallIndex, kSub).cons
        If curWinCons = useDefault Then
          curWinCons = iDefault.windCons
        End If
        strIDF "Construction", windowLayers(curWinCons - kindOfList(listWindow).firstChoice).nm
        strIDF "Base Surface", pExtWall(wallIndex).nm(floorIndex)
        strIDF "Outside Face Environment", ""
        numIDF "View Factor", 0.5
        strIDF "Shading Control", ""
        strIDF "Window Frame", ""
        'numIDF "Multiplier", iWindow(floorIndex, wallIndex, kSub).count
        numIDF "Multiplier", 1
        numIDF "Number of Surface Vertex Groups", 4
        numIDF "Vertex 1 X-coordinate", subLeftX
        numIDF "Vertex 1 Y-coordinate", subLeftY
        numIDF "Vertex 1 Z-coordinate", curFloorHeight + sillHeight + curSubHeight
        numIDF "Vertex 2 X-coordinate", subLeftX
        numIDF "Vertex 2 Y-coordinate", subLeftY
        numIDF "Vertex 2 Z-coordinate", curFloorHeight + sillHeight
        numIDF "Vertex 3 X-coordinate", subRightX
        numIDF "Vertex 3 Y-coordinate", subRightY
        numIDF "Vertex 3 Z-coordinate", curFloorHeight + sillHeight
        numIDF "Vertex 4 X-coordinate", subRightX
        numIDF "Vertex 4 Y-coordinate", subRightY
        numIDF "Vertex 4 Z-coordinate", curFloorHeight + sillHeight + curSubHeight, True
      Next mCnt
    End If
  Next kSub
  'define doors
  For kSub = 1 To doorsPerWall
    If iDoor(floorIndex, wallIndex, kSub).count > 0 Then
      curSubWidth = iDoor(floorIndex, wallIndex, kSub).widthSI
      If iDoor(floorIndex, wallIndex, kSub).width = useNumericDefault Then curSubWidth = iDefault.doorWidthSI
      curSubHeight = iDoor(floorIndex, wallIndex, kSub).heightSI
      If iDoor(floorIndex, wallIndex, kSub).height = useNumericDefault Then curSubHeight = iDefault.doorHeightSI
      For mCnt = 1 To iDoor(floorIndex, wallIndex, kSub).count
        'define the right edge of the door position than the left edge
        subRightX = curL * xSlope + x1
        subRightY = curL * ySlope + y1
        curL = curL + curSubWidth 'add the width of the door
        subLeftX = curL * xSlope + x1
        subLeftY = curL * ySlope + y1
        curL = curL + gapBetweenSubsurf 'add the gap bewteen door
        If iBuilding.epVersion >= epVersion300 Then
          objIDF "FenestrationSurface:Detailed"
        Else
          objIDF "SURFACE:HEATTRANSFER:SUB"
        End If
        strIDF "Name", iDoor(floorIndex, wallIndex, kSub).nm & "_" & Trim(Str(mCnt))
        strIDF "Type", "Door"
        strIDF "Construction", "DoorConstruction"
        strIDF "Base Surface", pExtWall(wallIndex).nm(floorIndex)
        strIDF "Outside Face Environment", ""
        numIDF "View Factor", 0.5
        strIDF "Shading Control", ""
        strIDF "Window Frame", ""
        numIDF "Multiplier", 1   'no multplier for doors
        numIDF "Number of Surface Vertex Groups", 4
        numIDF "Vertex 1 X-coordinate", subLeftX
        numIDF "Vertex 1 Y-coordinate", subLeftY
        numIDF "Vertex 1 Z-coordinate", curFloorHeight + curSubHeight
        numIDF "Vertex 2 X-coordinate", subLeftX
        numIDF "Vertex 2 Y-coordinate", subLeftY
        numIDF "Vertex 2 Z-coordinate", curFloorHeight
        numIDF "Vertex 3 X-coordinate", subRightX
        numIDF "Vertex 3 Y-coordinate", subRightY
        numIDF "Vertex 3 Z-coordinate", curFloorHeight
        numIDF "Vertex 4 X-coordinate", subRightX
        numIDF "Vertex 4 Y-coordinate", subRightY
        numIDF "Vertex 4 Z-coordinate", curFloorHeight + curSubHeight, True
      Next mCnt
    End If
  Next kSub
End If
End Sub

'------------------------------------------------------------------------
' Creates the portion of the IDF file that is related to horizontal
' surfaces such as roofs and floors including objects:
'    surface:heattransfer
' Also creates an attic zone if needed (if pitched roof)
'------------------------------------------------------------------------
Sub createIDFhorizSurfaces()
Dim isPitchedRoof As Boolean
'first see if pitched roof is to be developed or just flat roof
'the roof must have a height with roofPkHt to be pitched
'and roof corners and roof needs to be defined.
'first assume that it is flat
isPitchedRoof = False
If iBuilding.roofPkHt > 0 Then
  If numPRoof > 0 Then
    If numPRoofCorner > 0 Then
      isPitchedRoof = True
    End If
  End If
End If
If isPitchedRoof Then
  Call makePitchedRoofIDF
  Call makeAtticZoneIDF
  Call makeAtticFloorIDF
Else
  Call makeFlatRoofIDF
End If
Call makeIntermediateFloorsIDF
Call makeBottomFloorIDF
End Sub

'------------------------------------------------------------------------
' Use the definition from the roof
'   Roof is made up of a surface connecting normal corners with roof
'   corners (roof corners are negative indices).  The roof corners
'   are elevated to the roof peak height while normal corners are just
'   at the height of the attic floor.  This provides a sloped roof.
'------------------------------------------------------------------------
Sub makePitchedRoofIDF()
Dim heightOfRoofEdge As Single
Dim heightOfRoofPeak As Single
Dim iRoof As Integer
Dim kCorner As Integer
Dim jFloor As Integer
Dim iZone As Integer
Dim highFloor As Integer
'first locate the top floor that is active
highFloor = -1
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    highFloor = jFloor
  End If
Next jFloor
If highFloor = -1 Then
  MsgBox "An active floor was not found in makePitchedRoofIDF routine", vbExclamation, "Error"
  Exit Sub
End If
If iFloorPlan(highFloor).flr2flr = useNumericDefault Then
  heightOfRoofEdge = iFloorPlan(highFloor).heightOfFloor + iDefault.flr2flrSI
Else
  heightOfRoofEdge = iFloorPlan(highFloor).heightOfFloor + iFloorPlan(highFloor).flr2flrSI
End If
heightOfRoofPeak = heightOfRoofEdge + iBuilding.roofPkHtSI
For iRoof = 1 To numPRoof
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", pRoof(iRoof).nm
  strIDF "Surface Type", "Roof"
  strIDF "Construction Name", constInsulCombo(roofConsInsul).nm
  strIDF "Inside Face Environment", "AtticZone"
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Outside Face Environment", "Outdoors"
  Else
    strIDF "Outside Face Environment", "ExteriorEnvironment"
  End If
  strIDF "Outside Face Environment Object", ""
  strIDF "Sun Exposure", "SunExposed"
  strIDF "Wind Exposure", "WindExposed"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pRoof(iRoof).numRoofCrnrs)
  For kCorner = pRoof(iRoof).numRoofCrnrs To 1 Step -1
    If pRoof(iRoof).crnrs(kCorner) > 0 Then
      'if normal corner use normal edge of roof height (the eave)
      numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pRoof(iRoof).crnrs(kCorner)).xTransSI
      numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pRoof(iRoof).crnrs(kCorner)).yTransSI
      If kCorner > 1 Then
        numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfRoofEdge
      Else
        numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfRoofEdge, True
      End If
    Else
      'if roof corner then use roof peak height
      numIDF "Vertex X-coordinate " & Str(kCorner), pRoofCorner(-pRoof(iRoof).crnrs(kCorner)).xTransSI
      numIDF "Vertex Y-coordinate " & Str(kCorner), pRoofCorner(-pRoof(iRoof).crnrs(kCorner)).yTransSI
      If kCorner > 1 Then
        numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfRoofPeak
      Else
        numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfRoofPeak, True
      End If
    End If
  Next kCorner
Next iRoof
End Sub

'------------------------------------------------------------------------
' Define the zone for the attic. A single zone is for the entire attic.
'------------------------------------------------------------------------
Sub makeAtticZoneIDF()
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Zone"
Else
  objIDF "ZONE"
End If
strIDF "Name", "AtticZone"
numIDF "Relative North (not used)", 0
numIDF "X coord (not used)", 0
numIDF "Y coord (not used)", 0
numIDF "Z coord (not used)", 0
numIDF "Zone type (not used)", 1
numIDF "Multiplier", 1
strIDF "Ceiling Height", "" 'leave this blank to fix warning on ceiling height bug from 0.3 release (was iBuilding.roofPkHtSI)
numIDF "Volume (calculate)", 0
strIDF "Zone Inside Convection Algorithm", "Detailed", True
End Sub

'------------------------------------------------------------------------
' Make the attic floor from both zones (top floor zone and the attic
' perspective)
'------------------------------------------------------------------------
Sub makeAtticFloorIDF()
Dim highFloor As Integer
Dim jFloor As Integer
Dim iZone As Integer
Dim kCorner As Integer
Dim curZoneName As String
Dim floorHeight As Single
'first locate the top floor that is active
highFloor = -1
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    highFloor = jFloor
  End If
Next jFloor
If highFloor = -1 Then
  MsgBox "An active floor was not found in makeAtticFloorIDF routine", vbExclamation, "Error"
  Exit Sub
End If
If iFloorPlan(highFloor).flr2flr = useNumericDefault Then
  floorHeight = iFloorPlan(highFloor).heightOfFloor + iDefault.flr2flrSI
Else
  floorHeight = iFloorPlan(highFloor).heightOfFloor + iFloorPlan(highFloor).flr2flrSI
End If
'go through the zone and make a surface for each
'looking up
For iZone = 1 To numPZone
  curZoneName = pZone(iZone).nm(highFloor)
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", "AtticFloor_" & curZoneName
  strIDF "Surface Type", "Ceiling"
  strIDF "Construction Name", "IntFloorConstruction"
  strIDF "Inside Face Environment", pZone(iZone).nm(highFloor)
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Outside Face Environment", "Surface"
  Else
    strIDF "Outside Face Environment", "OtherZone"
  End If
  strIDF "Outside Face Environment Object", "AtticFloor_" & curZoneName & "_Mirror"
  strIDF "Sun Exposure", "NoSun"
  strIDF "Wind Exposure", "NoWind"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pZone(iZone).numZoneCrnrs)
  For kCorner = pZone(iZone).numZoneCrnrs To 1 Step -1
    numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).xTransSI
    numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).yTransSI
    If kCorner > 1 Then
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight
    Else
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight, True
    End If
  Next kCorner
  'now make the mirror copy of the surface for heat transfer from the other direction
  'looking down
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", "AtticFloor_" & curZoneName & "_Mirror"
  strIDF "Surface Type", "Floor"
  strIDF "Construction Name", "IntFloorConstruction"
  strIDF "Inside Face Environment", "AtticZone"
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Outside Face Environment", "Surface"
  Else
    strIDF "Outside Face Environment", "OtherZone"
  End If
  strIDF "Outside Face Environment Object", "AtticFloor_" & curZoneName
  strIDF "Sun Exposure", "NoSun"
  strIDF "Wind Exposure", "NoWind"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pZone(iZone).numZoneCrnrs)
  For kCorner = 1 To pZone(iZone).numZoneCrnrs
    numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).xTransSI
    numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).yTransSI
    If kCorner < pZone(iZone).numZoneCrnrs Then
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight
    Else
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight, True
    End If
  Next kCorner
Next iZone
End Sub

'------------------------------------------------------------------------
' Make intermediate floors if necessary
'------------------------------------------------------------------------
Sub makeIntermediateFloorsIDF()
Dim intFloor As Integer
Dim jFloor As Integer
Dim iZone As Integer
Dim kCorner As Integer
Dim curZoneName As String
Dim heightOfSlab As Single
Dim numActiveFloors
'first locate the intermediate floor that is active
intFloor = -1
numActiveFloors = 0
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    intFloor = jFloor
    numActiveFloors = numActiveFloors + 1
  End If
Next jFloor
If intFloor = -1 Then
  MsgBox "An active floor was not found in makeIntermediateFloorIDF routine", vbExclamation, "Error"
  Exit Sub
End If
'if not more than two active floors then cannot have intermediate floor
If numActiveFloors < 2 Then Exit Sub
'       1 = basement
'       2 = lower
'       3 = middle (or all)
'       4 = top
If iFloorPlan(1).active And iFloorPlan(2).active Then
  Call makeAnIntermediateFloor(1, 2)
ElseIf iFloorPlan(1).active And iFloorPlan(3).active Then
  Call makeAnIntermediateFloor(1, 3)
ElseIf iFloorPlan(1).active And iFloorPlan(4).active Then
  Call makeAnIntermediateFloor(1, 4)
End If
If iFloorPlan(2).active And iFloorPlan(3).active Then
  Call makeAnIntermediateFloor(2, 3)
ElseIf iFloorPlan(2).active And iFloorPlan(4).active Then
  Call makeAnIntermediateFloor(2, 4)
End If
If iFloorPlan(3).active And iFloorPlan(4).active Then
  Call makeAnIntermediateFloor(3, 4)
End If
End Sub

'------------------------------------------------------------------------
' Make a specific intermediate floor (called by makeIntermediateFloorsIDF)
'------------------------------------------------------------------------
Sub makeAnIntermediateFloor(lowFlr As Integer, hiFlr As Integer)
Dim iZone As Integer
Dim kCorner As Integer
Dim lowZoneName As String
Dim hiZoneName As String
Dim floorHeight As Single
floorHeight = iFloorPlan(hiFlr).heightOfFloorSI
'go through the zone and make a surface for each
For iZone = 1 To numPZone
  lowZoneName = pZone(iZone).nm(lowFlr)
  hiZoneName = pZone(iZone).nm(hiFlr)
  ' looking up
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", "Floor_" & lowZoneName & "_" & hiZoneName
  strIDF "Surface Type", "Ceiling"
  strIDF "Construction Name", "IntFloorConstruction"
  strIDF "Inside Face Environment", lowZoneName
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Outside Face Environment", "Surface"
  Else
    strIDF "Outside Face Environment", "OtherZoneSurface"
  End If
  strIDF "Outside Face Environment Object", "Floor_" & hiZoneName & "_" & lowZoneName
  strIDF "Sun Exposure", "NoSun"
  strIDF "Wind Exposure", "NoWind"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pZone(iZone).numZoneCrnrs)
  ' this from viewed from above should be going through the corners counter clockwise
  For kCorner = pZone(iZone).numZoneCrnrs To 1 Step -1
    numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).xTransSI
    numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).yTransSI
    If kCorner > 1 Then
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight
    Else
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight, True
    End If
  Next kCorner
  'now make the mirror copy of the surface for heat transfer from the other direction
  'looking up the points should be clockwise
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", "Floor_" & hiZoneName & "_" & lowZoneName
  strIDF "Surface Type", "Floor"
  strIDF "Construction Name", "IntFloorConstruction"
  strIDF "Inside Face Environment", hiZoneName
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Outside Face Environment", "Surface"
  Else
    strIDF "Outside Face Environment", "OtherZoneSurface"
  End If
  strIDF "Outside Face Environment Object", "Floor_" & lowZoneName & "_" & hiZoneName
  strIDF "Sun Exposure", "NoSun"
  strIDF "Wind Exposure", "NoWind"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pZone(iZone).numZoneCrnrs)
  ' this from viewed underneath should be going through the corners counter clockwise
  ' which is the same as clockwise when viewed from above.
  For kCorner = 1 To pZone(iZone).numZoneCrnrs
    numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).xTransSI
    numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).yTransSI
    If kCorner < pZone(iZone).numZoneCrnrs Then
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight
    Else
      numIDF "Vertex Z-coordinate " & Str(kCorner), floorHeight, True
    End If
  Next kCorner
Next iZone
End Sub

'------------------------------------------------------------------------
' Make a slab or foundation or simlpy bottom floor using the zone corners
'------------------------------------------------------------------------
Sub makeBottomFloorIDF()
Dim botFloor As Integer
Dim jFloor As Integer
Dim iZone As Integer
Dim kCorner As Integer
Dim curZoneName As String
Dim heightOfSlab As Single
'first locate the bottom floor that is active
botFloor = -1
For jFloor = maxNumFloorPlan To 1 Step -1
  If iFloorPlan(jFloor).active Then
    botFloor = jFloor
  End If
Next jFloor
If botFloor = -1 Then
  MsgBox "An active floor was not found in makeBottomFloorIDF routine", vbExclamation, "Error"
  Exit Sub
End If
heightOfSlab = iFloorPlan(botFloor).heightOfFloorSI
'go through the zone and make a surface for each
For iZone = 1 To numPZone
  curZoneName = pZone(iZone).nm(botFloor)
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", "Foundation_" & curZoneName
  strIDF "Surface Type", "Floor"
  strIDF "Construction Name", "BotFloorConstruction"
  strIDF "Inside Face Environment", pZone(iZone).nm(botFloor)
  strIDF "Outside Face Environment", "Ground"
  strIDF "Outside Face Environment Object", ""
  strIDF "Sun Exposure", "NoSun"
  strIDF "Wind Exposure", "NoWind"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pZone(iZone).numZoneCrnrs)
  ' this from viewed underneath should be going through the corners counter clockwise
  ' which is the same as clockwise when viewed from above.
  For kCorner = 1 To pZone(iZone).numZoneCrnrs
    numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).xTransSI
    numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).yTransSI
    If kCorner < pZone(iZone).numZoneCrnrs Then
      numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfSlab
    Else
      numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfSlab, True
    End If
  Next kCorner
Next iZone
End Sub

'------------------------------------------------------------------------
' Make a flat roof with the same pieces as the zones
'------------------------------------------------------------------------
Sub makeFlatRoofIDF()
Dim highFloor As Integer
Dim jFloor As Integer
Dim iZone As Integer
Dim kCorner As Integer
Dim curZoneName As String
Dim heightOfRoof As Single
'first locate the top floor that is active
highFloor = -1
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    highFloor = jFloor
  End If
Next jFloor
If highFloor = -1 Then
  MsgBox "An active floor was not found in makeFlatRoofIDF routine", vbExclamation, "Error"
  Exit Sub
End If
If iFloorPlan(highFloor).flr2flr = useNumericDefault Then
  heightOfRoof = iFloorPlan(highFloor).heightOfFloorSI + iDefault.flr2flrSI
Else
  heightOfRoof = iFloorPlan(highFloor).heightOfFloorSI + iFloorPlan(highFloor).flr2flrSI
End If
'go through the zone and make a surface for each
For iZone = 1 To numPZone
  curZoneName = pZone(iZone).nm(highFloor)
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "BuildingSurface:Detailed"
  Else
    objIDF "SURFACE:HEATTRANSFER"
  End If
  strIDF "Name", "RoofAbove_" & curZoneName
  strIDF "Surface Type", "Roof"
  strIDF "Construction Name", constInsulCombo(roofConsInsul).nm
  strIDF "Inside Face Environment", pZone(iZone).nm(highFloor)
  If iBuilding.epVersion >= epVersion300 Then
    strIDF "Outside Face Environment", "Outdoors"
  Else
    strIDF "Outside Face Environment", "ExteriorEnvironment"
  End If
  strIDF "Outside Face Environment Object", ""
  strIDF "Sun Exposure", "SunExposed"
  strIDF "Wind Exposure", "WindExposed"
  numIDF "View Factor to Ground", 0#
  numIDF "Number of Surface Vertex Groups", CSng(pZone(iZone).numZoneCrnrs)
  For kCorner = pZone(iZone).numZoneCrnrs To 1 Step -1
    numIDF "Vertex X-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).xTransSI
    numIDF "Vertex Y-coordinate " & Str(kCorner), pCorner(pZone(iZone).crnrs(kCorner)).yTransSI
    If kCorner > 1 Then
      numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfRoof
    Else
      numIDF "Vertex Z-coordinate " & Str(kCorner), heightOfRoof, True
    End If
  Next kCorner
Next iZone
End Sub

'------------------------------------------------------------------------
' Create all the constructions that will be used by the walls
' and roof
'------------------------------------------------------------------------
Sub createIDFconstructions()
Dim jFloor As Integer
Dim iWall As Integer
Dim kMat As Integer
Dim mCi As Integer
Dim i As Integer
Dim constructionPt As Integer
Dim insulArrayOffset As Integer
Dim curIns As Integer
Dim curMat As Integer
Dim Found As Integer
Dim concForIntFloor As Integer
Dim concForBotFloor As Integer

'clear the isUsed flags for:
'material, constLayer, insulation, windowGlassGas, windowLayers
For i = 1 To numMaterials
  MATERIAL(i).isUsed = False
Next i
For i = 1 To numConstLayer
  constLayer(i).isUsed = False
Next i
For i = 1 To numInsulation
  insulation(i).isUsed = False
Next i
For i = 1 To numWindowGlassGas
  windowGlassGas(i).isUsed = False
Next i
For i = 1 To numWindowLayers
  windowLayers(i).isUsed = False
Next i
Erase constInsulCombo
'define the insulation array offset based on the first choice for insulation
insulArrayOffset = kindOfList(listInsulation).firstChoice - 1
' define an ceiling air gap
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Material:AirGap"
Else
  objIDF "MATERIAL:AIR"
End If
strIDF "Name", "AirGap"
numIDF "Resistance", 0.1762, True
'intermediate floor constructions
Select Case iBuilding.floorCons - kindOfList(listFloorConstruction).firstChoice
  Case 0 '4in LW concrete
    concForIntFloor = 31
  Case 1 '6in LW concrete
    concForIntFloor = 32
  Case 2 '8in LW concrete
    concForIntFloor = 33
  Case 3 '4in HW concrete
    concForIntFloor = 22
  Case 4 '6in HW concrete
    concForIntFloor = 30
  Case 5 '8in HW concrete
    concForIntFloor = 27
End Select
'material and construction for internal floors
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Material"
Else
  objIDF "MATERIAL:REGULAR"
End If
strIDF "name", "ConcForIntFloor"
strIDF "Roughness", "Smooth"
numIDF "Thickness", MATERIAL(concForIntFloor).thick
numIDF "Conductivity", MATERIAL(concForIntFloor).conduct
numIDF "Density", MATERIAL(concForIntFloor).dens
numIDF "Specific Heat", MATERIAL(concForIntFloor).spheat
numIDF "Thermal Emittance", MATERIAL(concForIntFloor).emit
numIDF "Solar Absorptance", MATERIAL(concForIntFloor).solAbs
numIDF "Visible Absorptance", MATERIAL(concForIntFloor).visAbs, True
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Construction"
Else
  objIDF "CONSTRUCTION"
End If
strIDF "Name", "IntFloorConstruction"
strIDF "Concrete", "ConcForIntFloor", True
'Material and Construction and InternalMass
'Using 4" thick plywood as "typical of all furniture and mass in a space
'including paper (which is "thicker" when in file drawers)
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Material"
Else
  objIDF "MATERIAL:REGULAR"
End If
strIDF "Using Plywood", "FurnitureMaterial"
strIDF "Roughness", "MediumSmooth"
numIDF "Thickness", 0.1 '3.281 * 0.1 * 12 = 3.94" thick - about 4" thick
numIDF "Conductivity", 11
numIDF "Density", 544.62
numIDF "Specific Heat", 1210
numIDF "Thermal Emittance", 0.9
numIDF "Solar Absorptance", 0.78
numIDF "Visible Absorptance", 0.78, True
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Construction"
Else
  objIDF "CONSTRUCTION"
End If
strIDF "Name", "FurnitureConstruction"
strIDF "Material", "FurnitureMaterial", True
'bottom floor constructions
Select Case iBuilding.botFloorCons - kindOfList(listFloorConstruction).firstChoice
  Case 0 '4in LW concrete
    concForBotFloor = 31
  Case 1 '6in LW concrete
    concForBotFloor = 32
  Case 2 '8in LW concrete
    concForBotFloor = 33
  Case 3 '4in HW concrete
    concForBotFloor = 22
  Case 4 '6in HW concrete
    concForBotFloor = 30
  Case 5 '8in HW concrete
    concForBotFloor = 27
End Select
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Material"
Else
  objIDF "MATERIAL:REGULAR"
End If
strIDF "name", "ConcForBotFloor"
strIDF "Roughness", "Smooth"
numIDF "Thickness", MATERIAL(concForBotFloor).thick
numIDF "Conductivity", MATERIAL(concForBotFloor).conduct
numIDF "Density", MATERIAL(concForBotFloor).dens
numIDF "Specific Heat", MATERIAL(concForBotFloor).spheat
numIDF "Thermal Emittance", MATERIAL(concForBotFloor).emit
numIDF "Solar Absorptance", MATERIAL(concForBotFloor).solAbs
numIDF "Visible Absorptance", MATERIAL(concForBotFloor).visAbs, True
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Construction"
Else
  objIDF "CONSTRUCTION"
End If
strIDF "Name", "BotFloorConstruction"
strIDF "Concrete", "ConcForBotFloor", True
'DOOR construction (material 13 HF-B8 Wood 2.5 in)
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Material"
Else
  objIDF "MATERIAL:REGULAR"
End If
strIDF "name", "WoodForDoor"
strIDF "Roughness", "Smooth"
numIDF "Thickness", MATERIAL(13).thick
numIDF "Conductivity", MATERIAL(13).conduct
numIDF "Density", MATERIAL(13).dens
numIDF "Specific Heat", MATERIAL(13).spheat
numIDF "Thermal Emittance", MATERIAL(13).emit
numIDF "Solar Absorptance", MATERIAL(13).solAbs
numIDF "Visible Absorptance", MATERIAL(13).visAbs, True
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Construction"
Else
  objIDF "CONSTRUCTION"
End If
strIDF "Name", "DoorConstruction"
strIDF "Material", "WoodForDoor", True
'interior vertical wall constructions
Select Case iBuilding.intWallCons - kindOfList(listIntWallCons).firstChoice
  Case 0  'gyp + gap + concrete + gap + gyp
    'gypson
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Material"
    Else
      objIDF "MATERIAL:REGULAR"
    End If
    strIDF "name", "GypForInterior"
    strIDF "Roughness", "Smooth"
    numIDF "Thickness", MATERIAL(gypForInterior).thick
    numIDF "Conductivity", MATERIAL(gypForInterior).conduct
    numIDF "Density", MATERIAL(gypForInterior).dens
    numIDF "Specific Heat", MATERIAL(gypForInterior).spheat
    numIDF "Thermal Emittance", MATERIAL(gypForInterior).emit
    numIDF "Solar Absorptance", MATERIAL(gypForInterior).solAbs
    numIDF "Visible Absorptance", MATERIAL(gypForInterior).visAbs, True
    'concrete block (LW 8in)
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Material"
    Else
      objIDF "MATERIAL:REGULAR"
    End If
    strIDF "name", "ConcForInterior"
    strIDF "Roughness", "Smooth"
    numIDF "Thickness", MATERIAL(concForInterior).thick
    numIDF "Conductivity", MATERIAL(concForInterior).conduct
    numIDF "Density", MATERIAL(concForInterior).dens
    numIDF "Specific Heat", MATERIAL(concForInterior).spheat
    numIDF "Thermal Emittance", MATERIAL(concForInterior).emit
    numIDF "Solar Absorptance", MATERIAL(concForInterior).solAbs
    numIDF "Visible Absorptance", MATERIAL(concForInterior).visAbs, True
    'construction
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Construction"
    Else
      objIDF "CONSTRUCTION"
    End If
    strIDF "Name", "IntWallConstruction"
    strIDF "Gypsum", "GypForInterior"
    strIDF "AirGap", "AirGap"
    strIDF "Concrete", "ConcForInterior"
    strIDF "AirGap", "AirGap"
    strIDF "Gypsum", "GypForInterior", True
  Case 1  'concrete
    'concrete block (LW 8in)
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Material"
    Else
      objIDF "MATERIAL:REGULAR"
    End If
    strIDF "name", "ConcForInterior"
    strIDF "Roughness", "Smooth"
    numIDF "Thickness", MATERIAL(concForInterior).thick
    numIDF "Conductivity", MATERIAL(concForInterior).conduct
    numIDF "Density", MATERIAL(concForInterior).dens
    numIDF "Specific Heat", MATERIAL(concForInterior).spheat
    numIDF "Thermal Emittance", MATERIAL(concForInterior).emit
    numIDF "Solar Absorptance", MATERIAL(concForInterior).solAbs
    numIDF "Visible Absorptance", MATERIAL(concForInterior).visAbs, True
    'construction
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Construction"
    Else
      objIDF "CONSTRUCTION"
    End If
    strIDF "Name", "IntWallConstruction"
    strIDF "Concrete", "ConcForInterior", True
  Case 2  'gyp + wood frame + gyp
    'gypson
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Material"
    Else
      objIDF "MATERIAL:REGULAR"
    End If
    strIDF "name", "GypForInterior"
    strIDF "Roughness", "Smooth"
    numIDF "Thickness", MATERIAL(gypForInterior).thick
    numIDF "Conductivity", MATERIAL(gypForInterior).conduct
    numIDF "Density", MATERIAL(gypForInterior).dens
    numIDF "Specific Heat", MATERIAL(gypForInterior).spheat
    numIDF "Thermal Emittance", MATERIAL(gypForInterior).emit
    numIDF "Solar Absorptance", MATERIAL(gypForInterior).solAbs
    numIDF "Visible Absorptance", MATERIAL(gypForInterior).visAbs, True
    'construction
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Construction"
    Else
      objIDF "CONSTRUCTION"
    End If
    strIDF "Name", "IntWallConstruction"
    strIDF "Gypsum", "GypForInterior"
    strIDF "AirGap", "AirGap"
    strIDF "Gypsum", "GypForInterior", True
End Select
' define all of the materials that are used
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    For iWall = 1 To numPExtWall
      If pExtWall(iWall).cons(jFloor) = useDefault Then
        constructionPt = iDefault.extWallCons
      Else
        constructionPt = pExtWall(iWall).cons(jFloor)
      End If
      For kMat = 1 To constLayer(constructionPt).matCount
        curMat = constLayer(constructionPt).matInd(kMat)
        Select Case curMat
          Case constLayerInsul
            If pExtWall(iWall).insul(jFloor) = useDefault Then
              curIns = iDefault.extWallInsul - insulArrayOffset
            Else
              curIns = pExtWall(iWall).insul(jFloor) - insulArrayOffset
            End If
            Call defineCurIns(curIns)
          Case constLayerAirGap
            'do nothing - an air gap is already defined
          Case Else
            Call defineCurMat(curMat)
        End Select
      Next kMat
      ' now actual make the construction
      ' only if the combination of insulation and construction has not yet occured
      pExtWall(iWall).consInsul(jFloor) = defineTheConstruction(constructionPt, curIns)
    Next iWall
  End If
Next jFloor
'now define the roof
constructionPt = iBuilding.roofCons
If Not constLayer(constructionPt).isUsed Then
  For kMat = 1 To constLayer(constructionPt).matCount
    curMat = constLayer(constructionPt).matInd(kMat)
    Select Case curMat
      Case constLayerInsul
        curIns = iBuilding.roofInsul - insulArrayOffset
        Call defineCurIns(curIns)
      Case constLayerAirGap
        'do nothing - an air gap is already defined
      Case Else
        Call defineCurMat(curMat)
    End Select
  Next kMat
  roofConsInsul = defineTheConstruction(constructionPt, curIns)
  constLayer(constructionPt).isUsed = True
End If
End Sub


'------------------------------------------------------------------------
' Create a material definition for an insulation level
'------------------------------------------------------------------------
Sub defineCurIns(insulationLevel As Integer)
If Not insulation(insulationLevel).isUsed Then
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Material:NoMass"
  Else
    objIDF "MATERIAL:REGULAR-R"
  End If
  strIDF "Name", "R-" & Trim(Str(insulation(insulationLevel).rValue))
  strIDF "Roughness", "Rough"
  numIDF "Resistance", insulation(insulationLevel).rValue / 5.682
  numIDF "Thermal Emittance", 0.9
  numIDF "Solar Absorptance", 0.75
  numIDF "Visible Absorptance", 0.75, True
  'now set the flag to true so this value won't be defined more than once
  insulation(insulationLevel).isUsed = True
End If
End Sub


'------------------------------------------------------------------------
' Create a material definition for opaque material
'------------------------------------------------------------------------
Sub defineCurMat(materialIndex As Integer)
If Not MATERIAL(materialIndex).isUsed Then
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Material"
  Else
    objIDF "MATERIAL:REGULAR"
  End If
  strIDF MATERIAL(materialIndex).desc, MATERIAL(materialIndex).nm
  strIDF "Roughness", "Smooth"
  numIDF "Thickness", MATERIAL(materialIndex).thick
  numIDF "Conductivity", MATERIAL(materialIndex).conduct
  numIDF "Density", MATERIAL(materialIndex).dens
  numIDF "Specific Heat", MATERIAL(materialIndex).spheat
  numIDF "Thermal Emittance", MATERIAL(materialIndex).emit
  numIDF "Solar Absorptance", MATERIAL(materialIndex).solAbs
  numIDF "Visible Absorptance", MATERIAL(materialIndex).visAbs, True
  MATERIAL(materialIndex).isUsed = True
End If
End Sub


'------------------------------------------------------------------------
' now actual make the construction
' only if the combination of insulation and construction has not yet occured
'------------------------------------------------------------------------
Function defineTheConstruction(constIndex As Integer, insIndex As Integer) As Integer
Dim mCi As Integer
Dim Found As Integer
Dim kMat As Integer
Dim curMat As Integer
Found = 0
For mCi = 1 To numConstInsulCombo
  If constInsulCombo(mCi).constLayerPt = constIndex Then
    If constInsulCombo(mCi).insulationPt = insIndex Then
      Found = mCi
      Exit For
    End If
  End If
Next mCi
If Found = 0 Then
  numConstInsulCombo = numConstInsulCombo + 1
  constInsulCombo(numConstInsulCombo).constLayerPt = constIndex
  constInsulCombo(numConstInsulCombo).insulationPt = insIndex
  constInsulCombo(numConstInsulCombo).nm = constLayer(constIndex).nm & "_R-" & Trim(Str(insulation(insIndex).rValue))
  defineTheConstruction = numConstInsulCombo
  If iBuilding.epVersion >= epVersion300 Then
    objIDF "Construction"
  Else
    objIDF "CONSTRUCTION"
  End If
  strIDF "Name", constInsulCombo(numConstInsulCombo).nm
  For kMat = 1 To constLayer(constIndex).matCount
    curMat = constLayer(constIndex).matInd(kMat)
    Select Case curMat
      Case constLayerInsul
        strIDF "Material", "R-" & Trim(Str(insulation(insIndex).rValue))
      Case constLayerAirGap
        strIDF "Air Gap", "AirGap"
      Case Else
        If kMat = constLayer(constIndex).matCount Then
          strIDF "Material", MATERIAL(curMat).nm, True
        Else
          strIDF "Material", MATERIAL(curMat).nm
        End If
    End Select
  Next kMat
Else 'already in the list
  defineTheConstruction = Found
End If
End Function


'------------------------------------------------------------------------
' Create the constructions that correspond to windows
'------------------------------------------------------------------------
Sub createIDFwindowLayers()
Dim jFloor As Integer
Dim iWall As Integer
Dim kWin As Integer
Dim mLay As Integer
Dim nGlassGas As Integer
Dim curCons As Integer
Dim curConsChoice As Integer
Dim Found As Integer
Dim curGlassGas As String
Dim i As Integer
Dim firstWindowChoice As Integer
firstWindowChoice = kindOfList(listWindow).firstChoice
'create names for the different constructions based on the windows 4.0 number
For i = firstWindowChoice To kindOfList(listWindow).lastChoice
  windowLayers(i - firstWindowChoice).isUsed = False
  windowLayers(i - firstWindowChoice).nm = "WinLib_" & Left(listOfChoices(i), 4)
  'Debug.Print i - firstWindowChoice, windowLayers(i - firstWindowChoice).nm
Next i
For nGlassGas = 1 To numWindowGlassGas
  windowGlassGas(nGlassGas).isUsed = False
Next nGlassGas
' define all of the materials that are used
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    For iWall = 1 To numPExtWall
      For kWin = 1 To windowsPerWall
        If iWindow(jFloor, iWall, kWin).count > 0 Then
          curConsChoice = iWindow(jFloor, iWall, kWin).cons
          If curConsChoice = useDefault Then  'default
            curConsChoice = iDefault.windCons
          End If
          curCons = curConsChoice - kindOfList(listWindow).firstChoice
          If Not windowLayers(curCons).isUsed Then
            For mLay = 1 To windowLayers(curCons).layerCount
              'search for the layer string
              Found = 0
              For nGlassGas = 1 To numWindowGlassGas
                If windowLayers(curCons).layerName(mLay) = windowGlassGas(nGlassGas).nm Then
                  Found = nGlassGas
                  Exit For
                End If
              Next nGlassGas
              If Found > 0 Then
                If Not windowGlassGas(Found).isUsed Then
                  curGlassGas = windowGlassGas(Found).prop
                  If iBuilding.epVersion >= epVersion300 Then
                    curGlassGas = Replace(curGlassGas, "MATERIAL:WindowGlass", "WindowMaterial:Glazing")
                    curGlassGas = Replace(curGlassGas, "MATERIAL:WindowGas", "WindowMaterial:Gas")
                  End If
                  'include the material description in the IDF file
                  Print #idfFileHandle,
                  Print #idfFileHandle, curGlassGas
                  windowGlassGas(Found).isUsed = True
                End If
              Else
              End If
            Next mLay
            'now show the construction
            If iBuilding.epVersion >= epVersion300 Then
              objIDF "Construction"
            Else
              objIDF "CONSTRUCTION"
            End If
            strIDF listOfChoices(curConsChoice), windowLayers(curCons).nm
            For mLay = 1 To windowLayers(curCons).layerCount
              If mLay < windowLayers(curCons).layerCount Then
                strIDF "layer", windowLayers(curCons).layerName(mLay)
              Else
                strIDF "layer", windowLayers(curCons).layerName(mLay), True
              End If
            Next mLay
            windowLayers(curCons).isUsed = True
          End If
        End If
      Next kWin
    Next iWall
  End If
Next jFloor
End Sub


'------------------------------------------------------------------------
' routine creates the zone and all of the space gain objects that are
' based on styles
' also create internal mass for each space
'------------------------------------------------------------------------
Sub createIDFspacegains()
Dim curStyle As Integer
Dim curZoneName As String
Dim curZoneArea As Single
Dim jFloor As Integer
Dim iZone As Integer
For jFloor = 1 To maxNumFloorPlan
  If iFloorPlan(jFloor).active Then
    For iZone = 1 To numPZone
      curZoneName = pZone(iZone).nm(jFloor)
      curStyle = pZone(iZone).style(jFloor)
      If curStyle = useDefault Then
        curStyle = iDefault.style
      End If
      curStyle = curStyle - kindOfList(listStyle).firstChoice + 1
      curZoneArea = pZone(iZone).areaSI
      iStyle(curStyle).isUsed = True
      'define the zone
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "Zone"
      Else
        objIDF "ZONE"
      End If
      strIDF "Name", pZone(iZone).nm(jFloor)
      numIDF "Relative North (not used)", 0
      numIDF "X coord (not used)", 0
      numIDF "Y coord (not used)", 0
      numIDF "Z coord (not used)", 0
      numIDF "Zone type (not used)", 1
      numIDF "Multiplier", iFloorPlan(jFloor).numFlr
      If iFloorPlan(jFloor).flr2flr = useNumericDefault Then
        numIDF "Ceiling Height", iDefault.flr2flrSI
      Else
        numIDF "Ceiling Height", iFloorPlan(jFloor).flr2flrSI
      End If
      numIDF "Volume (calculate)", 0
      strIDF "Zone Inside Convection Algorithm", "Detailed", True
      'people
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "People"
      Else
        objIDF "PEOPLE"
      End If
      If iBuilding.epVersion <> epVersion121 And iBuilding.epVersion <> epVersion122 Then
        strIDF "Name", curZoneName & "_Peop"
      End If
      strIDF "Zone", curZoneName
      If iBuilding.epVersion >= epVersion220 Then
        strIDF "Schedule", "People_" & iStyle(curStyle).nm & "_Sch"
        strIDF "Number of People calculation method", "area/person"
        strIDF "Number of People", ""
        strIDF "People per Zone Area", ""
        numIDF "Zone area per person", iStyle(curStyle).peopDensUseSI
        numIDF "Frac Radiant", 0.2
        strIDF "user specified sensible fraction", "autocalculate"
      Else
        numIDF "Number", curZoneArea / iStyle(curStyle).peopDensUseSI
        strIDF "Schedule", "People_" & iStyle(curStyle).nm & "_Sch"
        numIDF "Frac Radiant", 0.2
      End If
      strIDF "Activity Schedule", "Activity_" & iStyle(curStyle).nm & "_Sch", True
      'lights
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "Lights"
      Else
        objIDF "LIGHTS"
      End If
      If iBuilding.epVersion <> epVersion121 And iBuilding.epVersion <> epVersion122 Then
        strIDF "Name", curZoneName & "_Lite"
      End If
      strIDF "Zone", curZoneName
      strIDF "Schedule", "Lights_" & iStyle(curStyle).nm & "_Sch"
      If iBuilding.epVersion >= epVersion220 Then
        strIDF "Design Level calculation method", "Watts/area"
        strIDF "Lighting Level", ""
        numIDF "Watts per Zone Area", iStyle(curStyle).liteDensUseSI
        strIDF "Watts per Person", ""
      Else
        numIDF "Design Level", curZoneArea * iStyle(curStyle).liteDensUseSI
      End If
      numIDF "Return Air Fraction", 0
      numIDF "Radiant Fraction", 0.37
      numIDF "Fraction Visible", 0.18
      numIDF "Fraction Replaceable", 0
      strIDF "Light End Use Key", "GeneralLights", True
      'electrical equipment
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "ElectricEquipment"
      Else
        objIDF "ELECTRIC EQUIPMENT"
      End If
      If iBuilding.epVersion <> epVersion121 And iBuilding.epVersion <> epVersion122 Then
        strIDF "Name", curZoneName & "_Elec"
      End If
      strIDF "Zone", curZoneName
      strIDF "Schedule", "Elec_" & iStyle(curStyle).nm & "_Sch"
      If iBuilding.epVersion >= epVersion220 Then
        strIDF "Design Level calculation method", "Watts/area"
        strIDF "Design Level", ""
        numIDF "Watts per Zone Area", iStyle(curStyle).eqpDensUseSI
        strIDF "Watts per Person", ""
      Else
        numIDF "Design Level", curZoneArea * iStyle(curStyle).eqpDensUseSI
      End If
      numIDF "Fraction Latent", 0
      numIDF "Fraction Radiant", 0.25
      numIDF "Fraction Lost", 0
      numIDF "End Use Category", 0, True
      'interal mass for furntiture
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "InternalMass"
      Else
        objIDF "SURFACE:HEATTRANSFER:INTERNALMASS"
      End If
      strIDF "Name", curZoneName & "_Furniture1"
      strIDF "Construction", "FurnitureConstruction"
      strIDF "Zone", curZoneName
      ' since not changing the density or thickness
      ' can only change the area so prorate by density times thickness
      strIDF "Area", 0.5 * curZoneArea * iStyle(curStyle).furnDensSI / (544.62 * 0.1), True
      If iBuilding.epVersion >= epVersion300 Then
        objIDF "InternalMass"
      Else
        objIDF "SURFACE:HEATTRANSFER:INTERNALMASS"
      End If
      strIDF "Name", curZoneName & "_Furniture2"
      strIDF "Construction", "FurnitureConstruction"
      strIDF "Zone", curZoneName
      ' since not changing the density or thickness
      ' can only change the area so prorate by density times thickness
      strIDF "Area", 0.5 * curZoneArea * iStyle(curStyle).furnDensSI / (544.62 * 0.1), True
    Next iZone
  End If
Next jFloor
End Sub

'------------------------------------------------------------------------
' create schedule objects for each of the schedules specified in
' the zones:
'      "People_" & iStyle(curStyle).nm & "_Sch"
'      "Activity_" & iStyle(curStyle).nm & "_Sch"
'      "Lights_" & iStyle(curStyle).nm & "_Sch"
'      "Elec_" & iStyle(curStyle).nm & "_Sch"
'------------------------------------------------------------------------
Sub createIDFschedules()
Dim jStyle As Integer
Dim timeRangeString As String
Dim wdTimeRangeStart As Integer 'weekdays
Dim wdTimeRangeEnd As Integer
Dim satTimeRangeStart As Integer 'saturdays
Dim satTimeRangeEnd As Integer
Dim sunTimeRangeStart As Integer 'sundays
Dim sunTimeRangeEnd As Integer
Dim curStyleName As String
Dim peopleOffRatio As Single
Dim liteOffRatio As Single
Dim elecOffRatio As Single
For jStyle = 1 To numIStyle
  If iStyle(jStyle).isUsed Then
    'weekdays
    timeRangeString = listOfChoices(iStyle(jStyle).weekdayTimeRange)
    Call parseTimeRange(timeRangeString, wdTimeRangeStart, wdTimeRangeEnd)
    'saturday
    timeRangeString = listOfChoices(iStyle(jStyle).weekdayTimeRange)
    Call parseTimeRange(timeRangeString, satTimeRangeStart, satTimeRangeEnd)
    'sunday
    timeRangeString = listOfChoices(iStyle(jStyle).weekdayTimeRange)
    Call parseTimeRange(timeRangeString, sunTimeRangeStart, sunTimeRangeEnd)
    'start to define the schedules
    curStyleName = iStyle(jStyle).nm
    'People
    If iStyle(jStyle).peopDensNonUseSI > 0 Then
      peopleOffRatio = iStyle(jStyle).peopDensUseSI / iStyle(jStyle).peopDensNonUseSI
    Else
      peopleOffRatio = 0
    End If
    Call makeCompactSchedule("People_" & curStyleName & "_Sch", wdTimeRangeStart, wdTimeRangeEnd, _
                   satTimeRangeStart, satTimeRangeEnd, sunTimeRangeStart, sunTimeRangeEnd, 1, peopleOffRatio)
    'Lights
    If iStyle(jStyle).liteDensUseSI > 0 Then
      liteOffRatio = iStyle(jStyle).liteDensNonUseSI / iStyle(jStyle).liteDensUseSI
    Else
      liteOffRatio = 0
    End If
    Call makeCompactSchedule("Lights_" & curStyleName & "_Sch", wdTimeRangeStart, wdTimeRangeEnd, _
                   satTimeRangeStart, satTimeRangeEnd, sunTimeRangeStart, sunTimeRangeEnd, 1, liteOffRatio)
    'Elec
    If iStyle(jStyle).eqpDensUseSI > 0 Then
      elecOffRatio = iStyle(jStyle).eqpDensNonUseSI / iStyle(jStyle).eqpDensUseSI
    Else
      elecOffRatio = 0
    End If
    Call makeCompactSchedule("Elec_" & curStyleName & "_Sch", wdTimeRangeStart, wdTimeRangeEnd, _
                   satTimeRangeStart, satTimeRangeEnd, sunTimeRangeStart, sunTimeRangeEnd, 1, elecOffRatio)
    'Activity (same value all year)
    If iBuilding.epVersion >= epVersion300 Then
      objIDF "Schedule:Compact"
    Else
      objIDF "SCHEDULE:COMPACT"
    End If
    strIDF "Name", "Activity_" & curStyleName & "_Sch"
    strIDF "Schedule Type", "AnyNumber"
    strIDF "Through", "Through: 12/31"
    strIDF "For", "For: AllDays"
    strIDF "Until", "Until: 24:00"
    numIDF "Value", 65, True
  End If
Next jStyle
End Sub

'------------------------------------------------------------------------
' Creates the SCHEDULE:COMPACT object for a specific case of
' start and end times expressed as 1 to 24 for weekdays, saturdays
' and sundays with constant values for on and off times.
'------------------------------------------------------------------------
Sub makeCompactSchedule(schName As String, wdStart As Integer, wdEnd As Integer, _
                        satStart As Integer, satEnd As Integer, _
                        sunStart As Integer, sunEnd As Integer, onVal As Single, offVal As Single)
If iBuilding.epVersion >= epVersion300 Then
  objIDF "Schedule:Compact"
Else
  objIDF "SCHEDULE:COMPACT"
End If
strIDF "Name", schName
strIDF "Schedule Type", "Fraction"
strIDF "Through", "Through: 12/31"
strIDF "For", "For: Weekdays"
If wdStart > 1 Then
  strIDF "Until", "Until: " & toTimeFormat(wdStart)
  numIDF "Value", offVal
End If
strIDF "Until", "Until: " & toTimeFormat(wdEnd)
numIDF "Value", onVal
If wdEnd < 24 Then
  strIDF "Until", "Until: 24:00"
  numIDF "Value", offVal
End If
strIDF "For", "For: Saturday"
If satStart > 1 Then
  strIDF "Until", "Until: " & toTimeFormat(satStart)
  numIDF "Value", offVal
End If
strIDF "Until", "Until: " & toTimeFormat(satEnd)
numIDF "Value", onVal
If satEnd < 24 Then
  strIDF "Until", "Until: 24:00"
  numIDF "Value", offVal
End If
strIDF "For", "For: Sunday AllOtherDays"
If sunStart > 1 Then
  strIDF "Until", "Until: " & toTimeFormat(sunStart)
  numIDF "Value", offVal
End If
If sunEnd < 24 Then
  strIDF "Until", "Until: " & toTimeFormat(sunEnd)
  numIDF "Value", onVal
  strIDF "Until", "Until: 24:00"
  numIDF "Value", offVal, True
Else
  strIDF "Until", "Until: 24"
  numIDF "Value", onVal, True
End If
End Sub

'------------------------------------------------------------------------
' Takes the string used in list of choices and parses it to find
' the starting and ending time
'------------------------------------------------------------------------
Sub parseTimeRange(trString As String, trStart As Integer, trEnd As Integer)
Dim toLoc As Integer
If trString = "All Hours" Then
  trStart = 1
  trEnd = 24
Else
  toLoc = InStr(trString, "to")
  trStart = getTimeNumber(Mid(trString, 7, 4))
  trEnd = getTimeNumber(Mid(trString, toLoc + 2))
End If
End Sub

'------------------------------------------------------------------------
' Converts an hour into hour:00 format
'------------------------------------------------------------------------
Function toTimeFormat(hr As Integer) As String
toTimeFormat = Trim(Str(hr)) & ":00"
End Function

'------------------------------------------------------------------------
' Convert a string that contains a time in the form of 8pm or noon into
' the 24 clock time of 8 or 12
'------------------------------------------------------------------------
Function getTimeNumber(timeString As String) As Integer
Dim trimmedTime As String
trimmedTime = LCase(Trim(timeString))
If trimmedTime = "noon" Then
  getTimeNumber = 12
Else
  If InStr(trimmedTime, "pm") > 0 Then
    getTimeNumber = 12 + Val(trimmedTime)
  Else
    getTimeNumber = Val(trimmedTime)
  End If
End If
End Function

'------------------------------------------------------------------------
' routine that prints to the IDF file the name of the object
' this is abstracted to a simple procedure so that the formatting
' can be easily changed
'------------------------------------------------------------------------
Sub objIDF(objectName As String, Optional lastItem As Boolean = False)
Print #idfFileHandle,
Print #idfFileHandle, objectName;
If lastItem Then
  Print #idfFileHandle, ";"
Else
  Print #idfFileHandle, ","
End If
End Sub

'------------------------------------------------------------------------
' routine that prints to the IDF file the value of a field and the
' field name (as a commment) this is for string values
' this is abstracted to a simple procedure so that the formatting
' can be easily changed
'------------------------------------------------------------------------
Sub strIDF(fieldName As String, fieldString As String, Optional lastItem As Boolean = False)
Print #idfFileHandle, "  "; fieldString;
If lastItem Then
  Print #idfFileHandle, ";";
Else
  Print #idfFileHandle, ",";
End If
Print #idfFileHandle, Tab(30); "!- "; fieldName
End Sub

'------------------------------------------------------------------------
' routine that prints to the IDF file the value of a field and the
' field name (as a commment) this is for numeric values
' this is abstracted to a simple procedure so that the formatting
' can be easily changed
'------------------------------------------------------------------------
Sub numIDF(fieldName As String, fieldValue As Single, Optional lastItem As Boolean = False)
Print #idfFileHandle, " "; fieldValue;
If lastItem Then
  Print #idfFileHandle, ";";
Else
  Print #idfFileHandle, ",";
End If
Print #idfFileHandle, Tab(30); "!- "; fieldName
End Sub

'------------------------------------------------------------------------
' Because the templates only list walls and zones and do not associate
' them together and EnergyPlus needs to know how they are related,
' this routine goes through each exterior and interior wall and looks
' for the zone(s) that it is associated with. It calls wallIsOnZone
' for each one and does some checking.
'------------------------------------------------------------------------
Sub associateWallsWithZones()
Dim iWall As Integer
Dim zoneOne As Integer
Dim zoneTwo As Integer
'exterior walls
For iWall = 1 To numPExtWall
  Call wallIsOnZone(pExtWall(iWall).startCorner, pExtWall(iWall).endCorner, zoneOne, zoneTwo)
  If zoneOne = 0 Or zoneTwo <> 0 Then
    MsgBox "Exterior wall: " & pExtWall(iWall).nm(1) & " is not associated with only one zones.", vbExclamation, "Template error"
  End If
  pExtWall(iWall).zone1 = zoneOne
Next iWall
'interior walls
For iWall = 1 To numPIntWall
  Call wallIsOnZone(pIntWall(iWall).startCorner, pIntWall(iWall).endCorner, zoneOne, zoneTwo)
  If zoneOne = 0 Or zoneTwo = 0 Then
    MsgBox "Interior wall: " & pIntWall(iWall).nm(1) & " is not associated with two zones." & _
    vbCrLf & pZone(zoneOne).nm(1) & " and " & pZone(zoneTwo).nm(1), vbExclamation, "Template error"
  End If
  pIntWall(iWall).zone1 = zoneOne
  pIntWall(iWall).zone2 = zoneTwo
Next iWall
End Sub

'------------------------------------------------------------------------
' Check what zone or zones the wall is part of.  The wall is defined
' as two corners (cornerStart and cornerEnd) and the results are returned
' are in zone1 and zone2. If it is an exterior wall then zone1 is the only
' one returned and zone2 is zero. If it is an interior wall then both
' zone1 and zone2 are returned as positive numbers.
'------------------------------------------------------------------------
Sub wallIsOnZone(cornerStart As Integer, cornerEnd As Integer, zone1 As Integer, zone2 As Integer)
Dim iZone As Integer
Dim jCorn As Integer
Dim zonesFound(2) As Integer
Dim otherCornerIndex As Integer
Dim zfi As Integer
zonesFound(1) = 0
zonesFound(2) = 0
zfi = 0
For iZone = 1 To numPZone
  For jCorn = 1 To pZone(iZone).numZoneCrnrs
    'check for the first corner
    If pZone(iZone).crnrs(jCorn) = cornerStart Then
      'if the first corner is found check for the second corner either
      'before or after the first corner
      'first next corner
      otherCornerIndex = jCorn + 1
      If otherCornerIndex = pZone(iZone).numZoneCrnrs + 1 Then otherCornerIndex = 1
      If cornerEnd = pZone(iZone).crnrs(otherCornerIndex) Then
        zfi = zfi + 1
        zonesFound(zfi) = iZone
        Exit For 'get out of loop for the corner
      Else
        'if that did not work check the previous corner instead
        otherCornerIndex = jCorn - 1
        If otherCornerIndex = 0 Then otherCornerIndex = pZone(iZone).numZoneCrnrs
        If cornerEnd = pZone(iZone).crnrs(otherCornerIndex) Then
          zfi = zfi + 1
          zonesFound(zfi) = iZone
          Exit For 'get out of loop for the corner
        End If
      End If
    End If
  Next jCorn
Next iZone
If zfi > 2 Then
  MsgBox "Wall found that is part of more then two zones.", vbInformation, "Error"
End If
zone1 = zonesFound(1)
zone2 = zonesFound(2) 'zero is returned if a second zone is not found (i.e. exterior walls)
End Sub

'------------------------------------------------------------------------
' This routine computes the SI values of all of the dimensions and other
' values that would be a different value when expressed in SI.
'------------------------------------------------------------------------
Sub convertToSI()
Dim ft2m As Single
Dim sqft2sqm As Single
Dim btuhPerSqft2WPerSqm As Single
Dim lbPerSqft2kgPerSqm As Single
Dim i As Integer, j As Integer, k As Integer
If newPlanInfo.isIPunits Then
  ft2m = 1 / 3.281
  sqft2sqm = 1 / (3.281 * 3.281)
  btuhPerSqft2WPerSqm = 1 / 0.316957210776545
  lbPerSqft2kgPerSqm = (3.281 * 3.281) / 2.2    '2.2 kg/lb and 3.281 ft/m
Else  'if in SI units then no computation is needed
  ft2m = 1
  sqft2sqm = 1
  btuhPerSqft2WPerSqm = 1
  lbPerSqft2kgPerSqm = 1
End If
For i = 1 To numPCorner
  pCorner(i).xTransSI = pCorner(i).xTrans * ft2m
  pCorner(i).yTransSI = pCorner(i).yTrans * ft2m
Next i
For i = 1 To numPRoofCorner
  pRoofCorner(i).xTransSI = pRoofCorner(i).xTrans * ft2m
  pRoofCorner(i).yTransSI = pRoofCorner(i).yTrans * ft2m
Next i
For i = 1 To numPExtWall
  pExtWall(i).lengthSI = pExtWall(i).length * ft2m
  For j = 1 To maxNumFloorPlan
    pExtWall(i).areaSI(j) = pExtWall(i).area(j) * sqft2sqm
  Next j
Next i
For i = 1 To numPZone
  pZone(i).areaSI = pZone(i).area * sqft2sqm
Next i
For i = 1 To numPRoof
  pRoof(i).areaSI = pRoof(i).area * sqft2sqm
Next i
iBuilding.roofPkHtSI = iBuilding.roofPkHt * ft2m
iDefault.flr2flrSI = iDefault.flr2flr * ft2m
iDefault.windWidthSI = iDefault.windWidth * ft2m
iDefault.windHeightSI = iDefault.windHeight * ft2m
iDefault.windOvrhngSI = iDefault.windOvrhng * ft2m
iDefault.windSetbckSI = iDefault.windSetbck * ft2m
iDefault.doorWidthSI = iDefault.doorWidth * ft2m
iDefault.doorHeightSI = iDefault.doorHeight * ft2m
For i = 1 To numIStyle
  iStyle(i).peopDensUseSI = iStyle(i).peopDensUse * sqft2sqm
  iStyle(i).peopDensNonUseSI = iStyle(i).peopDensNonUse * sqft2sqm
  iStyle(i).liteDensUseSI = iStyle(i).liteDensUse * btuhPerSqft2WPerSqm
  iStyle(i).liteDensNonUseSI = iStyle(i).liteDensNonUse * btuhPerSqft2WPerSqm
  iStyle(i).eqpDensUseSI = iStyle(i).eqpDensUse * btuhPerSqft2WPerSqm
  iStyle(i).eqpDensNonUseSI = iStyle(i).eqpDensNonUse * btuhPerSqft2WPerSqm
  iStyle(i).furnDensSI = iStyle(i).furnDens * lbPerSqft2kgPerSqm
Next i
For i = 1 To maxNumFloorPlan
  iFloorPlan(i).flr2flrSI = iFloorPlan(i).flr2flr * ft2m
  iFloorPlan(i).heightOfFloorSI = iFloorPlan(i).heightOfFloor * ft2m
Next i
For i = 1 To maxNumFloorPlan
  For j = 1 To numPExtWall
    For k = 1 To windowsPerWall
      iWindow(i, j, k).widthSI = iWindow(i, j, k).width * ft2m
      iWindow(i, j, k).heightSI = iWindow(i, j, k).height * ft2m
      iWindow(i, j, k).ovrhngSI = iWindow(i, j, k).ovrhng * ft2m
      iWindow(i, j, k).setbckSI = iWindow(i, j, k).setbck * ft2m
    Next k
  Next j
Next i
For i = 1 To maxNumFloorPlan
  For j = 1 To numPExtWall
    For k = 1 To doorsPerWall
      iDoor(i, j, k).widthSI = iDoor(i, j, k).width * ft2m
      iDoor(i, j, k).heightSI = iDoor(i, j, k).height * ft2m
    Next k
  Next j
Next i
End Sub

