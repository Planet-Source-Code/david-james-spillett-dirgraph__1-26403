Attribute VB_Name = "modNavigate"
Option Explicit

Private moBlocks As Collection

Public Sub NavInit()
    Set moBlocks = New Collection
End Sub

Public Sub NavAdd(ByVal iLeft As Long, ByVal iRight As Long, ByVal iTop As Long, ByVal iBottom As Long, sCaption As String, oDir As iDirObj)
Dim oNew As clsBlock
    Set oNew = New clsBlock
    oNew.Init iLeft, iRight, iTop, iBottom, sCaption, oDir
    moBlocks.Add oNew
End Sub

Public Function NavGetCaption(ByVal X As Long, ByVal Y As Long) As String
Dim oBlock As clsBlock
    NavGetCaption = ""
    If Not moBlocks Is Nothing Then
        For Each oBlock In moBlocks
            If oBlock.IsUnder(X, Y) Then NavGetCaption = oBlock.Caption
        Next
    End If
End Function

Public Function NavGetObject(ByVal X As Long, ByVal Y As Long) As iDirObj
Dim oBlock As clsBlock
    Set NavGetObject = Nothing
    If Not moBlocks Is Nothing Then
        For Each oBlock In moBlocks
            If oBlock.IsUnder(X, Y) Then Set NavGetObject = oBlock.Object
        Next
    End If
End Function

