VERSION 5.00
Begin VB.Form frmGraph 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   ControlBox      =   0   'False
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCaption 
      Interval        =   100
      Left            =   1740
      Top             =   3300
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1320
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private miOldWidth As Long
Private miOldHeight As Long
Private msOldAction As String

Private Sub Form_Load()
Const ciBorder As Long = 960
    Randomize Timer
    UpdateTitle
    Me.Move ciBorder, ciBorder, Screen.Width - ciBorder * 2, Screen.Height - ciBorder * 2
End Sub

Private Function UpdateTitle() As Boolean
Dim sCaption As String
    sCaption = "DirGraph v" & App.Major & "." & App.Minor & "." & App.Revision
    If gsCurAction <> "" Then sCaption = sCaption & ": " & gsCurAction
    If sCaption <> Me.Caption Then
        Me.Caption = sCaption
        UpdateTitle = True
    Else
        UpdateTitle = False
    End If
End Function

Private Sub Form_Resize()
    picGraph.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picGraph.ToolTipText = NavGetCaption(X, Y)
    'Debug.Print NavGetCaption(X, Y)
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oTemp As iDirObj
    Set oTemp = NavGetObject(X, Y)
    If oTemp Is Nothing Then
        gbForceRefresh = True
    Else
        If goTopObject.Path = oTemp.Path Then
            If Not goTopObject.Parent Is Nothing Then
                Debug.Print goTopObject.Class, goTopObject.Path
                Set goTopObject = goTopObject.Parent
                gbForceRefresh = True
            End If
        Else
            Set goTopObject = oTemp
            gbForceRefresh = True
        End If
    End If
    If Not goTopObject Is Nothing Then SaveSetting "DJS", App.Title, gcsRegLastDir, goTopObject.Path
End Sub

Private Sub tmrCaption_Timer()
    If frmControl.Visible = False And Me.WindowState <> 1 Then frmControl.Show 0, Me
    If Me.WindowState = 1 Then frmControl.Hide
    If UpdateTitle Or miOldHeight <> picGraph.Height Or miOldWidth <> picGraph.Width Or msOldAction <> gsCurAction Or gbForceRefresh Then
        gbForceRefresh = False
        miOldHeight = picGraph.Height
        miOldWidth = picGraph.Width
        msOldAction = gsCurAction
        If goTopObject Is Nothing Then
            ' nothing to draw
        Else
            If goTopWithTotal Is Nothing Then
                ' nothing to draw
                If goTopObject.TotalSize > 0 Then
                    picGraph.Cls
                    NavInit
                    DrawTree goTopObject, GetSetting("DJS", App.Title, gcsRegDisplayLevels, gcsDefLevels), picGraph, 0, GetSetting("DJS", App.Title, gcsRegLevelWidth, gcsDefLevelWidth), 0, picGraph.ScaleHeight - 1
                    If GetSetting("DJS", App.Title, gcsRegKey, gcsDefKey) <> 0 Then
                        DrawKey picGraph
                    End If
                End If
            Else
                ' draw as far as can
                If goTopWithTotal.TotalSize > 0 Then
                    picGraph.Cls
                    NavInit
                    DrawTree goTopWithTotal, GetSetting("DJS", App.Title, gcsRegDisplayLevels, gcsDefLevels), picGraph, 0, GetSetting("DJS", App.Title, gcsRegLevelWidth, gcsDefLevelWidth), 0, picGraph.ScaleHeight - 1
                    If GetSetting("DJS", App.Title, gcsRegKey, gcsDefKey) <> 0 Then
                        DrawKey picGraph
                    End If
                End If
            End If
        End If
    End If
End Sub
