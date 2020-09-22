VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iDirObj

Private moParent As iDirObj
Private moChildren As Collection
Private msPath As String

' Sizes: Currency used so we don't fall over at 2Gb
Private miOwnSize As Currency
Private miContentSize As Currency

' Flag to indicate 'async' operation [a few doevents here-or-there for screen update]
' [not yet implemented]
Private mbAyncScan As Boolean
Private mbAyncDraw As Boolean

Private moDateChange As clsChkDate
Private moDateAccess As clsChkDate


Private Sub Class_Initialize()
    mbAyncScan = False
    mbAyncDraw = False
    Set moDateChange = New clsChkDate
    Set moDateAccess = New clsChkDate
End Sub



Property Get iDirObj_Class() As geObjectType
    iDirObj_Class = geTypeDir
End Property



Property Get iDirObj_OwnSize() As Currency
    iDirObj_OwnSize = miOwnSize
End Property
Property Get iDirObj_ContentSize() As Currency
    iDirObj_ContentSize = miContentSize
End Property
Property Get iDirObj_TotalSize() As Currency
    iDirObj_TotalSize = miOwnSize + miContentSize
End Property

Property Get iDirObj_Parent() As iDirObj
Dim oFS As Object 'Scripting.FileSystemObject
Dim oDir As Object 'Scripting.Folder
    If moParent Is Nothing Then
        Set oFS = CreateObject("Scripting.FileSystemObject")
        Set oDir = oFS.GetFolder(goTopObject.Path)
        If Not oDir.ParentFolder Is Nothing Then
            Set goTopObject = NewObject(geTypeUnknown, oDir.ParentFolder, Nothing)
            goTopObject.PopulateTree
            Set goTopWithTotal = Nothing
            Set goTopMost = goTopObject
        End If
    End If
    Set iDirObj_Parent = moParent
End Property
Property Get iDirObj_Children() As Collection
    Set iDirObj_Children = moChildren
End Property

Property Get iDirObj_Path() As String
    iDirObj_Path = msPath
End Property
Property Get iDirObj_Name() As String
    If InStr(msPath, "\") = 0 Then
        iDirObj_Name = msPath
    Else
        iDirObj_Name = Mid(msPath, InStrRev(msPath, "\") + 1)
    End If
End Property



Public Sub iDirObj_Init(sPath As String, oParent As iDirObj)
    Set moParent = oParent
    msPath = Trim(Replace(sPath, "/", "\"))
    If Len(msPath) > 3 And Right(msPath, 1) = "\" Then msPath = Left(msPath, Len(msPath) - 1)
End Sub

Public Sub iDirObj_ParentLeaving()
    Set moParent = Nothing
End Sub

Private Property Get iDirObj_ToolTip() As String
    iDirObj_ToolTip = msPath & ": " & FormatSize(miOwnSize + miContentSize) & " (" & FormatSize(miOwnSize) & " in Files, " & FormatSize(miContentSize) & " in sub-dirs)"
End Property



'
' The meat of the class: Recursivly populates the tree under itself, calculating
' object sizes as it goes
'
' Called in three circumstances:
' 1. This is the first clsDir ever created
' 2. We want to refresh from this point down
' 3. We are moving up the tree, and need to calculate more
' [in case 3, oChild will already be populated so won't recalc to save time]
'
Public Sub iDirObj_PopulateTree(Optional oChild As iDirObj = Nothing, Optional oContainer As iDirObj = Nothing)
Dim oFS As Object 'Scripting.FileSystemObject
Dim oCurDir As Object 'Scripting.Folder
Dim oSubDir As Object 'Scripting.Folder
Dim oFile As Object 'Scripting.File
Dim oNewDir As iDirObj
Dim sOldAction As String
Dim oRealMe As iDirObj
Dim fT As Double
Dim bOK As Boolean

    If oContainer Is Nothing Then
        Set oRealMe = Me
    Else
        Set oRealMe = oContainer
    End If

    miContentSize = 0
    miOwnSize = 0
    Set moChildren = New Collection
    sOldAction = gsCurAction
    gsCurAction = msPath
    DoEvents
    
    Set oFS = CreateObject("Scripting.FileSystemObject")
    If oFS.FolderExists(msPath) Then ' sanity check
        bOK = True
        On Error Resume Next
        Set oCurDir = oFS.GetFolder(msPath)
        If Err.Number > 0 Then bOK = False
        On Error GoTo 0
        If bOK Then
            ' scan the filesystem ouselves, rather than using the Folder object's size
            ' method, as we need to record the sub-dir sizes for later [to save requeries]
            ' TODO?: this would probably be quicker if it didn't bother with the FSO, and
            ' just queried the directory structure itself [would save creating and destroying
            ' one object per file or folder]
            moDateAccess.ResetCount
            moDateChange.ResetCount
            For Each oSubDir In oCurDir.SubFolders
                Set oNewDir = NewObject(geTypeUnknown, oSubDir.Path, oRealMe)
                oNewDir.PopulateTree
                moChildren.Add oNewDir, oNewDir.Path
                miContentSize = miContentSize + oNewDir.TotalSize
                ' update 'TopWithTotal' so refresh knows what it can draw
                ' attempting to draw a dir that as yet has notgot a total [i.e. 'd:\' while a
                ' deep 'd:\program files' is being scanned] will cause devide by zero
                If goTopWithTotal Is Nothing Then Set goTopWithTotal = oRealMe
                If Len(goTopWithTotal.Path) > Len(msPath) Then
                    If InStr(goTopWithTotal.Path, ";") = 0 Then
                    Debug.Print Len(goTopWithTotal.Path), Len(msPath), goTopWithTotal.Path, msPath
                        Set goTopWithTotal = oRealMe
                    End If
                End If
                moDateAccess.AddDate oNewDir.MostRecentDateAccess
                moDateChange.AddDate oNewDir.MostRecentDateChange
            Next
            fT = Timer
            On Error Resume Next ' ignore errors. TODO: cleaner error handling
            For Each oFile In oCurDir.Files
                If Timer - fT > 0.2 Then
                    fT = Timer
                    gsCurAction = oFile.Path
                    DoEvents
                End If
                ' As we are scanning ourselves, keep the curdir's own size seperate from
                ' that of its contents, in case we want to display it.
                miOwnSize = miOwnSize + oFile.Size
                moDateChange.AddDate oFile.DateLastModified
                moDateAccess.AddDate oFile.DateLastAccessed
            Next
            On Error GoTo 0
        Else
            ' error getting directory: permission denied?
            moDateAccess.ResetCount
            moDateChange.ResetCount
        End If
    End If
    
    gsCurAction = sOldAction

End Sub



Property Get iDirObj_MostRecentDateAccess() As Date
    iDirObj_MostRecentDateAccess = moDateAccess.MostRecentDate
End Property
Property Get iDirObj_MostRecentDateChange() As Date
    iDirObj_MostRecentDateChange = moDateChange.MostRecentDate
End Property
Property Get iDirObj_Colour() As Long
    If GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour) = 1 Then
        iDirObj_Colour = moDateChange.Colour
    ElseIf GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour) = 2 Then
        iDirObj_Colour = moDateAccess.Colour
    End If
End Property

