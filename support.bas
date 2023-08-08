Attribute VB_Name = "support"
Option Explicit

'------------------------------------------------------------------------
' This type is defined to get information from
' the new plan dialog box
'------------------------------------------------------------------------
Private Type newPlanInfoType
  templateName As String
  numStories As Integer
  isFirstDiff As Boolean
  isTopDiff As Boolean
  isBasement As Boolean
  isIPunits As Boolean
  hvacSelect As Integer
End Type
Public newPlanInfo As newPlanInfoType

Public Const hvacSelPurch = 1
Public Const hvacSelDX = 2
Public Const hvacSelVAVair = 3
Public Const hvacSelVAVwater = 4

'------------------------------------------------------------------------
' Dates used when beta test version is shown
'------------------------------------------------------------------------

Public curDate As Variant
Public endDate As Variant

'------------------------------------------------------------------------
' Removes the path and drive and filename from the
' string and returns only the extension
'------------------------------------------------------------------------
Function extractExtension(wholePath As String) As String
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

