VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChkDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const mciColWhite As Long = &HFFFFFF
Const mciColYellow As Long = 65535 '&HFFFF
Const mciColGreen As Long = 65280 '&HFF00
Const mciColCyan As Long = &HFFFF00
Const mciColBlue As Long = &HFF0000
Const mciColMagenta As Long = &HFF00FF
Const mciColRed As Long = &HFF

Const mciColNone As Long = mciColWhite

Dim mdNewestDate As Date
Dim mbHasDate As Boolean

Private Sub Class_Initialize()
    ResetCount
End Sub

Public Sub ResetCount()
    mbHasDate = False
End Sub

Public Sub AddDate(dDate As Date)
    If Not mbHasDate Then mdNewestDate = dDate
    If dDate > mdNewestDate Then mdNewestDate = dDate
    mbHasDate = True
End Sub

Property Get MostRecentDate() As Date
    MostRecentDate = mdNewestDate
End Property

Property Get Colour() As Long
    If Not mbHasDate Then
        Colour = mciColNone
    ElseIf DateDiff("d", mdNewestDate, Now) < Val(GetSetting("DJS", App.Title, gcsRegDaysRoot & "1", 0)) Then
        Colour = GetSetting("DJS", App.Title, gcsRegColourRoot & "1", RGB(255, 255, 255))
    ElseIf DateDiff("d", mdNewestDate, Now) < Val(GetSetting("DJS", App.Title, gcsRegDaysRoot & "2", 0)) Then
        Colour = GetSetting("DJS", App.Title, gcsRegColourRoot & "2", RGB(255, 255, 255))
    ElseIf DateDiff("d", mdNewestDate, Now) < Val(GetSetting("DJS", App.Title, gcsRegDaysRoot & "3", 0)) Then
        Colour = GetSetting("DJS", App.Title, gcsRegColourRoot & "3", RGB(255, 255, 255))
    ElseIf DateDiff("d", mdNewestDate, Now) < Val(GetSetting("DJS", App.Title, gcsRegDaysRoot & "4", 0)) Then
        Colour = GetSetting("DJS", App.Title, gcsRegColourRoot & "4", RGB(255, 255, 255))
    ElseIf DateDiff("d", mdNewestDate, Now) < Val(GetSetting("DJS", App.Title, gcsRegDaysRoot & "5", 0)) Then
        Colour = GetSetting("DJS", App.Title, gcsRegColourRoot & "5", RGB(255, 255, 255))
    Else
        Colour = GetSetting("DJS", App.Title, gcsRegColourRoot & "6", RGB(255, 255, 255))
    End If
End Property


