VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private miLeft As Long
Private miRight As Long
Private miTop As Long
Private miBottom As Long
Private msCaption As String
Private moDir As iDirObj

Public Sub Init(iLeft As Long, iRight As Long, iTop As Long, iBottom As Long, sCaption As String, oDir As iDirObj)
    miLeft = iLeft
    miRight = iRight
    miTop = iTop
    miBottom = iBottom
    msCaption = sCaption
    Set moDir = oDir
End Sub

Public Function IsUnder(X As Long, Y As Long)
    IsUnder = False
    If miLeft <= X Then
        If miRight >= X Then
            If miTop <= Y Then
                If miBottom >= Y Then
                    IsUnder = True
                End If
            End If
        End If
    End If
End Function

Property Get Caption() As String
    Caption = msCaption
End Property
Property Get Object() As iDirObj
    Set Object = moDir
End Property

