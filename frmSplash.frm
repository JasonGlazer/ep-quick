VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7680
      Begin VB.Label Label1 
         Caption         =   "FREEWARE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         Caption         =   "Label1"
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   7455
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
lblProductName.Caption = App.Title
lblVersion.Caption = "Version: " & Format(mainForm.programVersion, "0.00")
lblDescription.Caption = App.Title & " is a program that helps create an EnergyPlus IDF file easily." _
& " With just a few choices the entire building description including surfaces and internal gains" _
& " may be described.  No HVAC systems are currently created with this version of the software." _
& vbCrLf & vbCrLf & "Copyright (c) 2003-2009 by Jason Glazer. All Rights reserved." _
& vbCrLf & "Portions include: clsMathParser copyright by Leonardo Volpi and VSFlexGrid copyright ComponentOne."
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

