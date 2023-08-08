VERSION 5.00
Begin VB.Form newPlan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbHVACselect 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CheckBox chkUseIP 
      Caption         =   "Use IP Units"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CheckBox chkTopFloorDiff 
      Caption         =   "Top Floor Different than Middle Floors"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CheckBox chkFirstFloorDiff 
      Caption         =   "First Floor Different than Middle Floors"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CheckBox chkBasement 
      Caption         =   "Basement"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ListBox lstOfTemplates 
      Height          =   6300
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.PictureBox pctTemplatePreview 
      Height          =   4095
      Left            =   5040
      ScaleHeight     =   4035
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "newPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------
' The exit routine for the dialog - unloads the form and copies
' values to public variables
'------------------------------------------------------------------------
Private Sub cmdContinue_Click()
If lstOfTemplates.Text = "" Then Exit Sub
' add warning if a polygon is selected
If InStr(lstOfTemplates.Text, "Poly") Then
  If MsgBox("Polygon templates require YOU to make sure all the corner locations " & _
  "are consistent. No checking to make sure the geometry makes sense is performed by " & _
  "the program. Templates that are not named Polygon are easier to use because the " & _
  "geometry is checked to ensure that it makes sense. " & vbCrLf & vbCrLf & _
  "Press OK to continue with this Polygon template. Press CANCEL to choose another " & _
  "template.", vbOKCancel, "WARNING") = vbCancel Then Exit Sub
End If
newPlanInfo.templateName = lstOfTemplates.Text
'newPlanInfo.numStories = Val(cmbNumStoriesAG.Text)
If chkFirstFloorDiff.Value = vbChecked Then
  newPlanInfo.isFirstDiff = True
Else
  newPlanInfo.isFirstDiff = False
End If
If chkTopFloorDiff.Value = vbChecked Then
  newPlanInfo.isTopDiff = True
Else
  newPlanInfo.isTopDiff = False
End If
If chkBasement.Value = vbChecked Then
  newPlanInfo.isBasement = True
Else
  newPlanInfo.isBasement = False
End If
If chkUseIP.Value = vbChecked Then
  newPlanInfo.isIPunits = True
Else
  newPlanInfo.isIPunits = False
End If
Select Case cmbHVACselect.Text
  Case "Zone-by-Zone Unitary"
    newPlanInfo.hvacSelect = hvacSelDX
  Case "Variable Air Volume - Air Cooled Chiller"
    newPlanInfo.hvacSelect = hvacSelVAVair
  Case "Variable Air Volume - Water Cooled Chiller"
    newPlanInfo.hvacSelect = hvacSelVAVwater
  Case "Air System Sizing"
    newPlanInfo.hvacSelect = hvacSelPurch
End Select
Unload Me
End Sub

'------------------------------------------------------------------------
' Initialize the form
'------------------------------------------------------------------------
Private Sub Form_Load()
cmbHVACselect.AddItem "Zone-by-Zone Unitary"
cmbHVACselect.AddItem "Variable Air Volume - Air Cooled Chiller"
cmbHVACselect.AddItem "Variable Air Volume - Water Cooled Chiller"
cmbHVACselect.AddItem "Air System Sizing"
cmbHVACselect.ListIndex = 1
Call getPlanNames
End Sub

'------------------------------------------------------------------------
' Get names of all PLN files
'
' A plan file contains a template for
' describing the layout plan for a floor
' of a building.
'------------------------------------------------------------------------
Sub getPlanNames()
Dim fs As FileSystemObject
Dim f As Folder
Dim fc As Files
Dim planFile As File
Dim s As String
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(App.Path & "\planTemplates")
Set fc = f.Files
For Each planFile In fc
  If extractExtension(planFile.name) = "PLN" Then
    lstOfTemplates.AddItem extractFileNameNoExt(planFile.name)
  End If
Next
End Sub

'------------------------------------------------------------------------
' When template is selected - display the WMF file
'------------------------------------------------------------------------
Private Sub lstOfTemplates_Click()
Dim locTemplate As String
Dim ln As Long
On Error Resume Next
locTemplate = App.Path & "\planTemplates\" & lstOfTemplates.Text & ".wmf"
Debug.Print lstOfTemplates.Text, locTemplate
ln = FileLen(locTemplate)
If Err.Number = 0 Then
  pctTemplatePreview.Picture = LoadPicture(locTemplate)
Else
  pctTemplatePreview.Picture = LoadPicture(App.Path & "\planTemplates\NoDrawingFound.wmf")
End If
On Error GoTo 0
End Sub
