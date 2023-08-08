VERSION 5.00
Begin VB.Form frmEULA 
   Caption         =   "End User License Agreement"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   8970
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEULA 
      BackColor       =   &H8000000F&
      Height          =   7095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmEULA.frx":0000
      Top             =   120
      Width           =   8775
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   7320
      Width           =   1095
   End
End
Attribute VB_Name = "frmEULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim eulaText As String
eulaText = "LICENSE AGREEMENT as of April 2007" & vbCrLf & vbCrLf
eulaText = eulaText & "Use of EP-Quick (the 'software') is governed by this "
eulaText = eulaText & "license agreement. By installing the software you agree to the terms "
eulaText = eulaText & "in this document. The software and its intellectual property are "
eulaText = eulaText & "protected by both United States Copyright Law and international copyright "
eulaText = eulaText & "treaties.  " & vbCrLf & vbCrLf
eulaText = eulaText & "This agreement is governed by the laws of the State of Illinois. Federal "
eulaText = eulaText & "and state courts sitting in Illinois shall have exclusive jurisdiction of "
eulaText = eulaText & "any disputes." & vbCrLf & vbCrLf
eulaText = eulaText & "The software is supplied 'as is' without warranty of any kind. Glazer "
eulaText = eulaText & "Software and Jason Glazer (1) disclaims any warranties, express or implied, including but "
eulaText = eulaText & "not limited to any implied warranties of merchantability, fitness for a "
eulaText = eulaText & "particular purpose, title or non-infringement, (2) does not assume any "
eulaText = eulaText & "legal liability or responsibility for the accuracy, completeness, or "
eulaText = eulaText & "usefulness of the software, (3) does not represent that the use of the "
eulaText = eulaText & "software would not infringe on privately owned rights, (4) does not warrant "
eulaText = eulaText & "that the software will function uninterrupted, that is it error-free or "
eulaText = eulaText & "that any errors will be corrected." & vbCrLf & vbCrLf
eulaText = eulaText & "In no event will Glazer Software or Jason Glazer be liable for any indirect, incidental, "
eulaText = eulaText & "consequential, special or punitive damages of any kind or nature, "
eulaText = eulaText & "including but not limited to loss of profits or loss of data, for any "
eulaText = eulaText & "reason whatsoever, whether such liability is asserted on the basis of a "
eulaText = eulaText & "contract, tort (including negligence or strict liability), or otherwise, "
eulaText = eulaText & "even if Glazer Software or Jason Glazerhas been warned of the possibility of such loss or "
eulaText = eulaText & "damages. In no even shall Glazer Software's or Jason Glazer's liability for damages arising "
eulaText = eulaText & "from or in connection with this agreement exceed the amount paid by you "
eulaText = eulaText & "for the software." & vbCrLf & vbCrLf
eulaText = eulaText & "Glazer Software" & vbCrLf
eulaText = eulaText & "www.glazersoftware.com"
txtEULA.Text = eulaText
End Sub
