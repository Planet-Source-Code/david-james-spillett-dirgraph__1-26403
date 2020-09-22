VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroup"
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



Private Sub Class_Initialize()
    mbAyncScan = False
    mbAyncDraw = False
End Sub



Property Get iDirObj_Class() As geObjectType
    iDirObj_Class = geTypeGroup
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
    Set iDirObj_Parent = Nothing
End Property
Property Get iDirObj_Children() As Collection
    Set iDirObj_Children = moChildren
End Property

Property Get iDirObj_Path() As String
    iDirObj_Path = msPath
End Property
Property Get iDirObj_Name() As String
    iDirObj_Name = msPath
End Property



Public Sub iDirObj_Init(sPath As String, oParent As iDirObj)
    Set moParent = oParent
    msPath = Trim(Replace(sPath, "/", "\"))
End Sub

Public Sub iDirObj_ParentLeaving()
    Set moParent = Nothing
End Sub

Private Property Get iDirObj_ToolTip() As String
    iDirObj_ToolTip = msPath & ": " & FormatSize(miContentSize)
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
'Dim oFS As Object 'Scripting.FileSystemObject
'Dim oCurDir As Object 'Scripting.Folder
'Dim oSubDir As Object 'Scripting.Folder
'Dim oFile As Object 'Scripting.File
Dim oNewDir As iDirObj
Dim sOldAction As String
Dim oRealMe As iDirObj
Dim asPath As Variant
Dim iLoop As Long

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
    
    asPath = Split(msPath, ";")
    
    For iLoop = LBound(asPath) To UBound(asPath)
        asPath(iLoop) = Trim(asPath(iLoop))
        Set oNewDir = NewObject(geTypeUnknown, asPath(iLoop), oRealMe)
        oNewDir.PopulateTree
        moChildren.Add oNewDir, oNewDir.Path
        miContentSize = miContentSize + oNewDir.TotalSize
        ' update 'TopWithTotal' so refresh knows what it can draw
        ' Assumes that clsGroups can only be top-level
        Set goTopWithTotal = oRealMe
    Next

    gsCurAction = sOldAction

End Sub



Property Get iDirObj_MostRecentDateAccess() As Date
    iDirObj_MostRecentDateAccess = DateSerial(1920, 1, 1)
End Property
Property Get iDirObj_MostRecentDateChange() As Date
    iDirObj_MostRecentDateChange = DateSerial(1920, 1, 1)
End Property
Property Get iDirObj_Colour() As Long
    iDirObj_Colour = &HFFFFFF
End Property

