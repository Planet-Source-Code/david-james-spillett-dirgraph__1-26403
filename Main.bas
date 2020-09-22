Attribute VB_Name = "modMain"
Option Explicit

Global goTopMost As iDirObj      ' what we are showing
Global goTopWithTotal As iDirObj ' what is done [for refresh during scan]
Global goTopObject As iDirObj    ' may not be either of the above
' goTopObject will always as at least as close to root as goWithTotal
' goWithTotal will always as at least as close to root as goTopMost
' Only goTopMost and foTopObject used ATM, goTopWithTotal is for Async usage

Global gsCurAction As String ' for progress reports
Global gbForceRefresh  As Boolean

Global Const gcsRegDisplayLevels As String = "Display Levels"
Global Const gcsRegLevelWidth As String = "Level Width"
Global Const gcsRegIncludeFiles As String = "Display Files"
Global Const gcsRegIncludeSmall As String = "Include Small Dir Summary"
Global Const gcsRegLastDir As String = "Most Recent"
Global Const gcsRegFreeSpace As String = "Free Space"
Global Const gcsRegColour As String = "Highlight"
Global Const gcsRegColourRec As String = "Highlight Recursive"
Global Const gcsRegColourRoot As String = "Colour "
Global Const gcsRegDaysRoot As String = "Days "
Global Const gcsRegWindowRoot As String = "Window "
Global Const gcsRegKey As String = "Show Key"

Global Const gcsDefLevels As Integer = 999
Global Const gcsDefLevelWidth As Integer = 100
Global Const gcsDefIncFiles As Integer = 1
Global Const gcsDefIncSmall As Integer = 1
Global Const gcsDefLastDir As String = "c:\"
Global Const gcsDefFreeSpace As Integer = 0
Global Const gcsDefColour As Integer = 1
Global Const gcsDefColourRec As Integer = 1
Global Const gcsDefKey As Integer = 1

Enum geObjectType
    geTypeUnknown = 0
    geTypeDir = 1
    geTypeDrive = 2
    geTypeNetShare = 3
    geTypeGroup = 4
End Enum
Public Function NewObject(iType As geObjectType, ByVal sPath As String, oParent As iDirObj)
Dim oTemp As iDirObj
Dim sTemp As String
    '
    If iType = geTypeUnknown Then
        If InStr(sPath, ";") <> 0 Then
            Set oTemp = New clsGroup
        ElseIf Len(Trim(sPath)) < 4 Then
            Set oTemp = New clsDrive
        ElseIf Left(Trim(sPath), 2) = "\\" Then
            sTemp = Mid(Trim(sPath), 3)
            If InStr(sTemp, "\") <> 0 Then sTemp = Mid(sTemp, InStr(sTemp, "\") + 1)
            If InStr(sTemp, "\") = 0 Then
                Set oTemp = New clsNetShare
            Else
                Set oTemp = New clsDir
            End If
        Else
            Set oTemp = New clsDir
        End If
    End If
    If iType = geTypeDir Then oTemp = New clsDir
    If iType = geTypeDrive Then oTemp = New clsDrive
    If iType = geTypeNetShare Then oTemp = New clsNetShare
    If iType = geTypeGroup Then oTemp = New clsGroup
    '
    oTemp.Init sPath, oParent
    Set NewObject = oTemp
    '
End Function
