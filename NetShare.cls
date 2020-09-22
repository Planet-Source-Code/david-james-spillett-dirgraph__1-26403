VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iDirObj
' clsDir is the general case
' Drives are considered special cases.

Private moDir As iDirObj
Private miFreeSpace As Currency
Private miTotalSize As Currency

Private Sub Class_Initialize()
    Set moDir = New clsDir
End Sub



Property Get iDirObj_Class() As geObjectType
    iDirObj_Class = geTypeNetShare
End Property



Property Get iDirObj_OwnSize() As Currency
    iDirObj_OwnSize = moDir.OwnSize
End Property
Property Get iDirObj_ContentSize() As Currency
    iDirObj_ContentSize = moDir.ContentSize
End Property
Property Get iDirObj_TotalSize() As Currency
    iDirObj_TotalSize = moDir.TotalSize + miFreeSpace
End Property

Property Get iDirObj_Parent() As iDirObj
' no parent for drives [until we support groups]
    Set iDirObj_Parent = Nothing
End Property
Property Get iDirObj_Children() As Collection
    Set iDirObj_Children = moDir.Children
End Property

Property Get iDirObj_Path() As String
    iDirObj_Path = moDir.Path
End Property
Property Get iDirObj_Name() As String
    iDirObj_Name = moDir.Name
End Property



Public Sub iDirObj_Init(sPath As String, oParent As iDirObj)
    moDir.Init sPath, oParent
End Sub

Public Sub iDirObj_ParentLeaving()
    moDir.ParentLeaving
End Sub

Private Property Get iDirObj_ToolTip() As String
    iDirObj_ToolTip = moDir.Path & ": " & FormatSize(moDir.TotalSize) & " [" & FormatSize(moDir.OwnSize) & " in Files, " & FormatSize(moDir.ContentSize) & " in sub-dirs, " & FormatSize(miFreeSpace) & " free (" & FormatSize(miTotalSize - miFreeSpace - moDir.TotalSize) & " unaccounted for)]"
End Property



Public Sub iDirObj_PopulateTree(Optional oChild As iDirObj = Nothing, Optional oContainer As iDirObj = Nothing)
Dim oFS As Object
Dim oDrv As Object
    moDir.PopulateTree oChild, Me
    If GetSetting("DJS", App.Title, gcsRegFreeSpace, gcsDefFreeSpace) > 1 Then
        On Error Resume Next
        Set oFS = CreateObject("scripting.filesystemobject")
        Set oDrv = oFS.getdrive(moDir.Path)
        miFreeSpace = oDrv.FreeSpace
        miTotalSize = oDrv.TotalSize
        On Error GoTo 0
    End If
End Sub



Property Get iDirObj_MostRecentDateAccess() As Date
    iDirObj_MostRecentDateAccess = moDir.MostRecentDateAccess
End Property
Property Get iDirObj_MostRecentDateChange() As Date
    iDirObj_MostRecentDateChange = moDir.MostRecentDateChange
End Property
Property Get iDirObj_Colour() As Long
    iDirObj_Colour = moDir.Colour
End Property

