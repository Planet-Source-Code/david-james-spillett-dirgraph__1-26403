VERSION 5.00
Begin VB.Form frmColour 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColour 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   667
      TabIndex        =   0
      Top             =   0
      Width           =   10035
   End
End
Attribute VB_Name = "frmColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim iR As Integer, iG As Integer, iB As Integer
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim lCol As Long
    For iR = 0 To 4
        For iG = 0 To 4
            For iB = 0 To 4
                x1 = picColour.ScaleWidth / 25 * (iR + iG * 5)
                x2 = picColour.ScaleWidth / 25 * (iR + 1 + iG * 5)
                y1 = picColour.ScaleHeight / 5 * iB
                y2 = picColour.ScaleHeight / 5 * (iB + 1)
                lCol = RGB(iR * 63, iG * 63, iB * 63)
                picColour.Line (x1, y1)-(x2, y2), lCol, BF
                If lCol = Val(Me.Tag) Then
                    If Val(Me.Tag) < 120 Then
                        picColour.Line (x1 + 3, y1 + 3)-(x2 - 3, y2 - 3), RGB(255, 255, 255), B
                    Else
                        picColour.Line (x1 + 3, y1 + 3)-(x2 - 3, y2 - 3), 0, B
                    End If
                End If
                
            Next
        Next
    Next
End Sub

Private Sub picColour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iR, iG, iB
    iB = Int(picColour.Point(X, Y) / 256 / 256)
    iR = picColour.Point(X, Y) - Int(picColour.Point(X, Y) / 256) * 256
    iG = (picColour.Point(X, Y) - iR - iB * 65536) / 256
    Me.Caption = "Colour: " & iR & "/" & iG & "/" & iB
End Sub

Private Sub picColour_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Tag = picColour.Point(X, Y)
    Me.Hide
End Sub
