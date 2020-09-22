Attribute VB_Name = "modDrawTree"
Option Explicit
Const micBackCol     As Long = &HE0E0E0
Const micDirCol      As Long = &HFFFFFF
Const micFileCol     As Long = &HF0F0F0
Const micLineCol     As Long = &H0
Const micFileTextCol As Long = &H666666

'
' Draw the given tree in the given picturebox
' iTop and iBottom are adhered to as limits
' iLeft and iRight are used for top level object only
' [will be moves iRight-iLeft to the right each level]
'
Public Sub DrawTree(oDir As iDirObj, iMaxDepth As Long, oPic As PictureBox, iLeft As Currency, iRight As Currency, iTop As Currency, iBottom As Currency)

Dim sName        As String
Dim sSize        As String
Dim bCanFit2Rows As Boolean
Dim bTooLong    As Boolean

Dim oSubDir       As iDirObj
Dim iSubLeft      As Currency
Dim iSubRight     As Currency
Dim iSubTop       As Currency
Dim iSubBottom    As Currency
Dim iSubDisplayed As Currency
Dim iSubSkipped   As Currency
Dim iSubSkippedF  As Currency
Dim oDateSkippedA As clsChkDate
Dim oDateSkippedC As clsChkDate
    
    ' draw own box
    oPic.Line (iLeft, iTop)-(iRight, iBottom), micDirCol, BF
    oPic.Line (iLeft, iTop)-(iRight, iBottom), micLineCol, B
    If GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour) > 0 And iBottom - iTop > 3 Then
        oPic.Line (iLeft + 2, iTop + 2)-(iRight - 2, iBottom - 2), oDir.Colour, B
    End If
    
    ' print name
    If iBottom - iTop > oPic.TextHeight(sName) * 2 + 2 Then
        oPic.CurrentY = iTop + (iBottom - iTop - oPic.TextHeight(sName) * 2) / 2
        bCanFit2Rows = True
    Else
        oPic.CurrentY = iTop + (iBottom - iTop - oPic.TextHeight(sName)) / 2
        bCanFit2Rows = False
    End If
    sName = oDir.Name
    bTooLong = False
    If Not bCanFit2Rows Then
        sName = sName & " " & FormatSize(oDir.TotalSize)
        Do While oPic.TextWidth(sName) > (iRight - iLeft - 8)
            sName = Right(sName, Len(sName) - 1)
            bTooLong = True
        Loop
        If bTooLong Then
            sName = "..." & sName
            Do While oPic.TextWidth(sName) > (iRight - iLeft - 8)
                sName = Left(sName, Len(sName) - 1)
                bTooLong = True
            Loop
        End If
    Else
        Do While oPic.TextWidth(sName) > (iRight - iLeft - 8)
            sName = Left(sName, Len(sName) - 1)
            bTooLong = True
        Loop
        If bTooLong Then
            sName = sName & "..."
            Do While oPic.TextWidth(sName) > (iRight - iLeft - 8)
                sName = Left(sName, Len(sName) - 4) & "..."
                bTooLong = True
            Loop
        End If
    End If
    oPic.ForeColor = micLineCol
    oPic.CurrentX = iLeft + (iRight - iLeft - oPic.TextWidth(sName)) / 2
    oPic.Print sName
    
    ' get size
    If bCanFit2Rows Then
        sSize = FormatSize(oDir.TotalSize)
        If oPic.TextWidth(sSize) > (iRight - iLeft - 8) Then
            sSize = "."
        End If
        oPic.CurrentX = iLeft + (iRight - iLeft - oPic.TextWidth(sSize)) / 2
        oPic.ForeColor = micLineCol
        oPic.Print sSize
    End If
    
    ' init
    iSubDisplayed = 0
    iSubSkipped = 0
    Set oDateSkippedA = New clsChkDate
    Set oDateSkippedC = New clsChkDate
    oDateSkippedA.ResetCount
    oDateSkippedC.ResetCount
    
    ' do subdirs and files
    If iMaxDepth > 0 Then
        iSubLeft = iLeft + (iRight - iLeft)
        iSubRight = iRight + (iRight - iLeft)
        For Each oSubDir In oDir.Children
            If oSubDir.TotalSize * (iBottom - iTop) / oDir.TotalSize > oPic.TextHeight("text") + 2 Then
                'Debug.Print "*" & oSubDir.Name, FormatSize(oSubDir.TotalSize), FormatSize(oDir.TotalSize), Int(CDbl(oSubDir.TotalSize) / CDbl(oDir.TotalSize) * (iBottom - iTop)), oPic.TextHeight("text") * 2
                ' display sub dir
                iSubTop = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
                iSubDisplayed = iSubDisplayed + oSubDir.TotalSize
                iSubBottom = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
                DrawTree oSubDir, iMaxDepth - 1, oPic, iSubLeft, iSubRight, iSubTop, iSubBottom
            Else
                'Debug.Print oSubDir.Name, FormatSize(oSubDir.TotalSize), FormatSize(oDir.TotalSize), Int(CDbl(oSubDir.TotalSize) / CDbl(oDir.TotalSize) * (iBottom - iTop)), oPic.TextHeight("text") * 2
                ' skip: is too small to fix text so will group with other smalls later
                iSubSkipped = iSubSkipped + oSubDir.TotalSize
                iSubSkippedF = iSubSkippedF + oSubDir.OwnSize
                oDateSkippedA.AddDate oSubDir.MostRecentDateAccess
                oDateSkippedC.AddDate oSubDir.MostRecentDateChange
            End If
        Next
        ' now a box for the skipped dirs, if 'small dirs summary' option is on
        If GetSetting("DJS", App.Title, gcsRegIncludeSmall, gcsDefIncSmall) <> 0 Then
            iSubTop = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
            iSubDisplayed = iSubDisplayed + iSubSkipped
            iSubBottom = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
            If iSubBottom - iSubTop > 1 Then
                oPic.Line (iSubLeft, iSubTop)-(iSubRight, iSubBottom), micDirCol, BF
                oPic.Line (iSubLeft, iSubTop)-(iSubRight, iSubBottom), micLineCol, B
                If GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour) > 0 And iSubBottom - iSubTop > 3 Then
                    If GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour) = 1 Then
                        oPic.Line (iSubLeft + 2, iSubTop + 2)-(iSubRight - 2, iSubBottom - 2), oDateSkippedC.Colour, B
                    ElseIf GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour) = 2 Then
                        oPic.Line (iSubLeft + 2, iSubTop + 2)-(iSubRight - 2, iSubBottom - 2), oDateSkippedA.Colour, B
                    End If
                End If
                NavAdd iSubLeft, iSubRight, iSubTop, iSubBottom, "Small dirs in " & oDir.Path & ": " & FormatSize(iSubSkipped) & " (total)", Nothing
                If iSubBottom - iSubTop > oPic.TextHeight("M") + 2 Then
                    sSize = "(" & FormatSize(iSubSkipped) & ")"
                    oPic.CurrentX = iSubLeft + (((iSubRight - iSubLeft) - oPic.TextWidth(sSize)) / 2)
                    oPic.CurrentY = iSubTop + (((iSubBottom - iSubTop) - oPic.TextHeight(sSize)) / 2)
                    oPic.ForeColor = micLineCol
                    oPic.Print sSize
                End If
                If iMaxDepth > 1 Then
                    ' if theres another level to go, and 'display files' option on, and
                    ' there were some files in the skipped dirs, show box
                    If iSubSkippedF > 0 And GetSetting("DJS", App.Title, gcsRegIncludeFiles, gcsDefIncFiles) Then
                        iSubDisplayed = iSubDisplayed - iSubSkipped
                        iSubTop = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
                        iSubDisplayed = iSubDisplayed + iSubSkippedF
                        iSubBottom = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
                        oPic.Line (iSubLeft + (iRight - iLeft), iSubTop)-(iSubRight + (iRight - iLeft), iSubBottom), micFileCol, BF
                        oPic.Line (iSubLeft + (iRight - iLeft), iSubTop)-(iSubRight + (iRight - iLeft), iSubBottom), micLineCol, B
                        NavAdd iSubLeft + (iRight - iLeft), iSubRight + (iRight - iLeft), iSubTop, iSubBottom, "Files in small dirs in " & oDir.Path & ": " & FormatSize(iSubSkippedF), Nothing
                        iSubDisplayed = iSubDisplayed - iSubSkippedF + iSubSkipped
                    End If
                End If
            End If
        End If
        ' and a box for files, if 'show files' option is on
        If GetSetting("DJS", App.Title, gcsRegIncludeFiles, gcsDefIncFiles) <> 0 Then
            iSubTop = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
            iSubDisplayed = iSubDisplayed + oDir.OwnSize
            iSubBottom = (iBottom - iTop) / oDir.TotalSize * iSubDisplayed + iTop
            If iSubBottom - iSubTop > 1 Then
                oPic.Line (iSubLeft, iSubTop)-(iSubRight, iSubBottom), micFileCol, BF
                oPic.Line (iSubLeft, iSubTop)-(iSubRight, iSubBottom), micLineCol, B
                NavAdd iSubLeft, iSubRight, iSubTop, iSubBottom, "Files in " & oDir.Path & ": " & FormatSize(oDir.OwnSize), Nothing
                If iSubBottom - iSubTop > oPic.TextHeight("M") + 2 Then
                    sSize = "(" & FormatSize(oDir.OwnSize) & ")"
                    oPic.CurrentX = iSubLeft + (((iSubRight - iSubLeft) - oPic.TextWidth(sSize)) / 2)
                    oPic.CurrentY = iSubTop + (((iSubBottom - iSubTop) - oPic.TextHeight(sSize)) / 2)
                    oPic.ForeColor = micFileTextCol
                    oPic.Print sSize
                End If
            End If
        End If
    End If

    ' add to navigation records
    NavAdd iLeft, iRight, iTop, iBottom, oDir.ToolTip, oDir
    
    ' refresh after each Dir printed
    ' removed: routine is so fast on these boxes that all the extra refeshes do is create flicker
    'oPic.Refresh

End Sub



Public Sub DrawKey(oPic As PictureBox)
Dim sAll As String
Dim iLoop As Long
Dim iTop As Long
Dim iLeft As Long
Dim iBottom As Long
Dim iRight As Long
    sAll = ""
    For iLoop = 1 To 6
        sAll = sAll & GetSetting("DJS", App.Title, gcsRegDaysRoot & iLoop, "++++") & "/"
    Next
    sAll = Left(sAll, Len(sAll) - 1)
    iLeft = oPic.ScaleWidth - oPic.TextWidth(sAll) - 5
    iTop = oPic.ScaleHeight - oPic.TextHeight(sAll) - 0
    iRight = oPic.ScaleWidth - 1
    iBottom = oPic.ScaleHeight - 1
    oPic.Line (iLeft, iTop)-(iRight, iBottom), micDirCol, BF
    oPic.Line (iLeft, iTop)-(iRight, iBottom), micLineCol, B
    oPic.CurrentX = iLeft + 2
    oPic.CurrentY = iTop

    For iLoop = 1 To 6
        oPic.ForeColor = GetSetting("DJS", App.Title, gcsRegColourRoot & iLoop, RGB(255, 255, 255))
        oPic.Print GetSetting("DJS", App.Title, gcsRegDaysRoot & iLoop, "++++");
        oPic.ForeColor = micLineCol
        If iLoop < 6 Then oPic.Print "/";
    Next
End Sub
