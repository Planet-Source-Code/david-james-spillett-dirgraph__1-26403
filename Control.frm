VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "1"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6195
   Icon            =   "Control.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   4920
      TabIndex        =   34
      Top             =   780
      Width           =   1215
   End
   Begin VB.CheckBox chkShowKey 
      Caption         =   "Show key on graph"
      Height          =   195
      Left            =   2820
      TabIndex        =   33
      Top             =   2940
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.ComboBox lstColour 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2580
      Width           =   2955
   End
   Begin VB.Frame frameColours 
      Caption         =   "Highlight Boundaries and Colours"
      Height          =   1035
      Left            =   60
      TabIndex        =   16
      Top             =   3300
      Width           =   3915
      Begin VB.CommandButton cmdDays 
         Caption         =   "Colour"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   28
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtDays 
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   27
         Text            =   "8888"
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmdDays 
         Caption         =   "Colour"
         Height          =   255
         Index           =   5
         Left            =   1860
         TabIndex        =   26
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtDays 
         Height          =   255
         Index           =   5
         Left            =   1380
         TabIndex        =   25
         Text            =   "8888"
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmdDays 
         Caption         =   "Colour"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   24
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtDays 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Text            =   "8888"
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmdDays 
         Appearance      =   0  'Flat
         Caption         =   "Colour"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   22
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDays 
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   21
         Text            =   "8888"
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdDays 
         Caption         =   "Colour"
         Height          =   255
         Index           =   2
         Left            =   1860
         TabIndex        =   20
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDays 
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   19
         Text            =   "8888"
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdDays 
         Caption         =   "Colour"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDays 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Text            =   "8888"
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.CheckBox chkColourRec 
      Caption         =   "Include child dir's files in calc"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   2940
      Value           =   1  'Checked
      Width           =   2835
   End
   Begin VB.ComboBox lstFreeSpace 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "Show Options"
      Height          =   255
      Left            =   3660
      TabIndex        =   12
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txtLevels 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply Options"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox chkSmallDirs 
      Caption         =   "Include blocks summerising dirs to small to display"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1920
      Value           =   1  'Checked
      Width           =   3915
   End
   Begin VB.CheckBox chkFiles 
      Caption         =   "Include blocks representing files"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   2220
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -60
      TabIndex        =   4
      Top             =   840
      Width           =   6435
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   1260
      Top             =   120
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change To:"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   1095
   End
   Begin VB.TextBox txtCur 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   60
      Width           =   4935
   End
   Begin VB.TextBox txtNew 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   420
      Width           =   4935
   End
   Begin VB.Frame Frame3 
      Height          =   2355
      Left            =   4800
      TabIndex        =   31
      Top             =   2580
      Width           =   1455
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "NOTE: Some options will not take effect until the tree is re-scanned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Highlight dirs"
      Height          =   255
      Left            =   60
      TabIndex        =   30
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Include free-space in size of drive bars"
      Height          =   195
      Left            =   2760
      TabIndex        =   13
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Label Label2 
      Caption         =   "Max Display Depth (levels):"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1500
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Level Width (pixels):"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label lblCur 
      Caption         =   "Scanning:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miOldWidth As Long
Private miOldHeight As Long

Private Sub cmdApply_Click()
Dim iLoop As Long
    txtLevels.Text = CInt(txtLevels.Text)
    If CInt(txtLevels.Text) < 1 Then txtLevels.Text = 1
    SaveSetting "DJS", App.Title, gcsRegDisplayLevels, txtLevels.Text
    txtWidth.Text = CInt(txtWidth.Text)
    If CInt(txtWidth.Text) < 50 Then txtWidth.Text = 50
    If CInt(txtWidth.Text) > 500 Then txtWidth.Text = 500
    SaveSetting "DJS", App.Title, gcsRegLevelWidth, txtWidth.Text
    SaveSetting "DJS", App.Title, gcsRegIncludeFiles, chkFiles.Value
    SaveSetting "DJS", App.Title, gcsRegIncludeSmall, chkSmallDirs.Value
    SaveSetting "DJS", App.Title, gcsRegFreeSpace, lstFreeSpace.ListIndex
    SaveSetting "DJS", App.Title, gcsRegColour, lstColour.ListIndex
    SaveSetting "DJS", App.Title, gcsRegColourRec, chkColourRec.Value
    For iLoop = 1 To 6
        SaveSetting "DJS", App.Title, gcsRegColourRoot & iLoop, txtDays(iLoop).BackColor
        SaveSetting "DJS", App.Title, gcsRegDaysRoot & iLoop, txtDays(iLoop).Text
    Next
    SaveSetting "DJS", App.Title, gcsRegKey, chkShowKey.Value
    gbForceRefresh = True
    cmdSize_Click
End Sub

Private Sub cmdChange_Click()
Dim oFS As Object 'Scripting.FileSystemObject
Dim asPath As Variant
Dim iLoop As Long
Dim bOK As Boolean
    Set oFS = CreateObject("Scripting.FileSystemObject")
    '
    txtNew.Text = Trim(txtNew.Text)
    If InStr(txtNew.Text, ";") = 0 Then
        If oFS.FolderExists(Trim(txtNew.Text)) Then
            bOK = True
        Else
            MsgBox "Specified Folder (" & txtNew.Text & ") does not exist"
            bOK = False
        End If
    Else
        asPath = Split(txtNew.Text, ";")
        bOK = True
        For iLoop = LBound(asPath) To UBound(asPath)
            If Not oFS.FolderExists(Trim(asPath(iLoop))) Then
                MsgBox "Specified Folder (" & Trim(asPath(iLoop)) & ") does not exist"
                bOK = False
            End If
        Next
    End If
    '
    If bOK Then
        SaveSetting "DJS", App.Title, gcsRegLastDir, txtNew.Text
        Set goTopWithTotal = Nothing
        Set goTopObject = NewObject(geTypeUnknown, txtNew.Text, Nothing)
        goTopObject.PopulateTree
        Set goTopWithTotal = Nothing
        Set goTopMost = goTopObject
        gbForceRefresh = True
    End If
    '
End Sub


Private Sub cmdDays_Click(Index As Integer)
    frmColour.Tag = txtDays(Index).BackColor
    frmColour.Show 1
    If frmColour.Tag <> "" Then txtDays(Index).BackColor = frmColour.Tag
End Sub

Private Sub cmdExit_Click()
    SaveSetting "DJS", App.Title, gcsRegWindowRoot & "Y", frmGraph.Top
    SaveSetting "DJS", App.Title, gcsRegWindowRoot & "X", frmGraph.Left
    SaveSetting "DJS", App.Title, gcsRegWindowRoot & "H", frmGraph.Height
    SaveSetting "DJS", App.Title, gcsRegWindowRoot & "W", frmGraph.Width
    End
End Sub

Private Sub cmdSize_Click()
Dim iSmall As Long
Dim iLarge As Long
    iSmall = cmdSize.Top + cmdSize.Height + (Me.Height - Me.ScaleHeight)
    iLarge = frameColours.Top + frameColours.Height + 60 + (Me.Height - Me.ScaleHeight)
    If Me.Height > iSmall Then
        Me.Height = iSmall
        cmdSize.Caption = "Show Options"
    Else
        If Me.Top + iLarge > Screen.Height Then Me.Top = Me.Top + iSmall - iLarge
        If Me.Top < 0 Then Me.Top = 0
        If Me.Top + iLarge > Screen.Height Then Me.Top = Screen.Height - Me.Height
        Me.Height = iLarge
        cmdSize.Caption = "Hide Options"
    End If
End Sub

Private Sub Form_Activate()
    If Command$ <> "" And goTopObject Is Nothing And Command$ = txtCur.Text Then
        cmdChange_Click
    End If
End Sub

Private Sub Form_Load()
Dim iLoop As Long
    If Not goTopObject Is Nothing Then txtCur.Text = goTopMost.Path
    If txtCur.Text = "" Then
        If Command$ <> "" Then
            txtNew = Command$
        Else
            txtNew.Text = GetSetting("DJS", App.Title, gcsRegLastDir, gcsDefLastDir)
        End If
    End If
    txtCur.Text = txtNew.Text
    txtLevels.Text = GetSetting("DJS", App.Title, gcsRegDisplayLevels, gcsDefLevels)
    txtWidth.Text = GetSetting("DJS", App.Title, gcsRegLevelWidth, gcsDefLevelWidth)
    chkFiles.Value = GetSetting("DJS", App.Title, gcsRegIncludeFiles, gcsDefIncFiles)
    chkSmallDirs.Value = GetSetting("DJS", App.Title, gcsRegIncludeSmall, gcsDefIncSmall)
    lstFreeSpace.AddItem "Never"
    lstFreeSpace.AddItem "For local drives only"
    lstFreeSpace.AddItem "For local and network drives"
    lstFreeSpace.ListIndex = GetSetting("DJS", App.Title, gcsRegFreeSpace, gcsDefFreeSpace)
    lstColour.AddItem "No"
    lstColour.AddItem "according to date of last file change"
    lstColour.AddItem "according to date of last file access"
    lstColour.ListIndex = GetSetting("DJS", App.Title, gcsRegColour, gcsDefColour)
    chkColourRec.Value = GetSetting("DJS", App.Title, gcsRegColourRec, gcsDefColourRec)
    If GetSetting("DJS", App.Title, gcsRegColourRoot & "1", "xxx") = "xxx" Then
        SaveSetting "DJS", App.Title, gcsRegColourRoot & "1", RGB(255, 255, 0) '&HFFFF
        SaveSetting "DJS", App.Title, gcsRegColourRoot & "2", RGB(0, 255, 0) '&HFF00
        SaveSetting "DJS", App.Title, gcsRegColourRoot & "3", RGB(0, 255, 255) '&HFFFF00
        SaveSetting "DJS", App.Title, gcsRegColourRoot & "4", RGB(0, 0, 255) '&HFF0000
        SaveSetting "DJS", App.Title, gcsRegColourRoot & "5", RGB(255, 0, 255) '&HFF00FF
        SaveSetting "DJS", App.Title, gcsRegColourRoot & "6", RGB(255, 0, 0) '&HFF
        SaveSetting "DJS", App.Title, gcsRegDaysRoot & "1", 1
        SaveSetting "DJS", App.Title, gcsRegDaysRoot & "2", 7
        SaveSetting "DJS", App.Title, gcsRegDaysRoot & "3", 30
        SaveSetting "DJS", App.Title, gcsRegDaysRoot & "4", 92
        SaveSetting "DJS", App.Title, gcsRegDaysRoot & "5", 365
    End If
    For iLoop = 1 To 6
        txtDays(iLoop).BackColor = GetSetting("DJS", App.Title, gcsRegColourRoot & iLoop, RGB(255, 255, 255))
        txtDays(iLoop).Text = GetSetting("DJS", App.Title, gcsRegDaysRoot & iLoop, "++++")
    Next
    chkShowKey = GetSetting("DJS", App.Title, gcsRegKey & "W", gcsDefKey)
    frmGraph.Top = GetSetting("DJS", App.Title, gcsRegWindowRoot & "Y", frmGraph.Top)
    frmGraph.Left = GetSetting("DJS", App.Title, gcsRegWindowRoot & "X", frmGraph.Left)
    frmGraph.Height = GetSetting("DJS", App.Title, gcsRegWindowRoot & "H", frmGraph.Height)
    frmGraph.Width = GetSetting("DJS", App.Title, gcsRegWindowRoot & "W", frmGraph.Width)
    UpdateTitle True
    cmdSize_Click
    Me.Top = frmGraph.Top + frmGraph.Height - Me.Height - (frmGraph.Height - frmGraph.ScaleHeight)
    Me.Left = frmGraph.Left + (frmGraph.Width - Me.Width) / 2
    If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height - 120
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width - 120
End Sub

Private Function UpdateTitle(Optional bForce As Boolean = False) As Boolean
Dim sCaption As String
    sCaption = "Controls"
    If gsCurAction <> "" Then sCaption = sCaption & ": " & gsCurAction
    If sCaption <> Me.Caption Then
        If bForce Then Me.Caption = sCaption
        UpdateTitle = True
    Else
        UpdateTitle = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
Const ciSmallWidth As Long = 1100
    If Me.Width > ciSmallWidth Then
        miOldHeight = Me.Height
        miOldWidth = Me.Width
        Me.Top = frmGraph.Top
        Me.Height = 50
        Me.Width = ciSmallWidth
        Me.Left = frmGraph.Left + frmGraph.Width - Me.Width
    Else
        Me.Width = miOldWidth
        Me.Height = miOldHeight
        If Me.Left + Me.Width > Screen.Width - 240 Then
            Me.Left = Screen.Width - Me.Left - 240
        End If
        If Me.Height + Me.Height > Screen.Height - 240 Then
            Me.Top = Screen.Height - Me.Height - 240
        End If
    End If
    Cancel = True
End Sub



Private Sub tmrUpdate_Timer()
    If UpdateTitle Then
        cmdChange.Enabled = False
    Else
        cmdChange.Enabled = True
    End If
    If gsCurAction <> "" Then
        txtCur = gsCurAction
        lblCur = "Scanning:"
    Else
        lblCur = "Showing:"
        If Not goTopObject Is Nothing Then
            txtCur = goTopObject.Path
        Else
            txtCur = "?"
        End If
    End If
    If Me.Height < 1000 Then
        Me.Top = frmGraph.Top
        Me.Left = frmGraph.Left + frmGraph.Width - Me.Width - 1200
    End If
End Sub

Private Sub txtNew_GotFocus()
    txtNew.SelStart = 0
    txtNew.SelLength = Len(txtNew.Text)
End Sub
