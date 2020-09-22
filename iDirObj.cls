VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iDirObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get Class() As geObjectType
End Property

Property Get OwnSize() As Currency
End Property
Property Get ContentSize() As Currency
End Property
Property Get TotalSize() As Currency
End Property

Property Get Parent() As iDirObj
End Property
Property Get Children() As Collection
End Property

Property Get Path() As String
End Property
Property Get Name() As String
End Property



Public Sub Init(sPath As String, oParent As iDirObj)
End Sub

Public Sub ParentLeaving()
End Sub

Property Get ToolTip() As String

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
Public Sub PopulateTree(Optional oChild As iDirObj = Nothing, Optional oContainer As iDirObj = Nothing)
End Sub



Property Get MostRecentDateAccess() As Date
End Property
Property Get MostRecentDateChange() As Date
End Property
Property Get Colour() As Long
End Property

