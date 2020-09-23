VERSION 5.00
Begin VB.Form frmPaperSizes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Paper Width"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5595
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLandscape 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Landscape"
      Height          =   255
      Left            =   1260
      TabIndex        =   7
      Top             =   600
      Width           =   1275
   End
   Begin VB.ComboBox cboUnits 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      ItemData        =   "frmPaperSizes.frx":0000
      Left            =   2970
      List            =   "frmPaperSizes.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   900
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4410
      TabIndex        =   5
      Top             =   630
      Width           =   1005
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   345
      Left            =   4410
      TabIndex        =   4
      Top             =   210
      Width           =   1005
   End
   Begin VB.TextBox txtWidth 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   900
      Width           =   1665
   End
   Begin VB.OptionButton optPaper 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custom:"
      Height          =   345
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   975
   End
   Begin VB.OptionButton optPaper 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preset:"
      Height          =   345
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.ComboBox cboPaper 
      Height          =   315
      ItemData        =   "frmPaperSizes.frx":002E
      Left            =   1260
      List            =   "frmPaperSizes.frx":0047
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   2955
   End
End
Attribute VB_Name = "frmPaperSizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngPaperWidth As Single
Dim smcScaleMode As ScaleModeConstants

Private Sub cboPaper_Click()
Me.ScaleMode = vbInches
smcScaleMode = vbInches
Select Case cboPaper.ListIndex
    Case 0
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 8.5
    Else
        lngPaperWidth = 11
    End If
    Case 1
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 11
    Else
        lngPaperWidth = 17
    End If
    Case 2
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 8.5
    Else
        lngPaperWidth = 14
    End If
    Case 3
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 23.4
    Else
        lngPaperWidth = 33.1
    End If
    Case 4
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 16.5
    Else
        lngPaperWidth = 23.4
    End If
    Case 5
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 11.7
    Else
        lngPaperWidth = 16.5
    End If
    Case 6
    If chkLandscape.Value = Unchecked Then
        lngPaperWidth = 8.3
    Else
        lngPaperWidth = 11.7
    End If
End Select
Me.ScaleMode = vbTwips
End Sub

Private Sub cboUnits_Click()
Select Case cboUnits.ListIndex
    Case 0
        smcScaleMode = vbInches
    Case 1
        smcScaleMode = vbPixels
    Case 2
        smcScaleMode = vbMillimeters
End Select
End Sub

Private Sub chkLandscape_Click()
cboPaper_Click
End Sub

Private Sub cmdApply_Click()
fMainForm.ActiveForm.ScaleMode = smcScaleMode
fMainForm.ActiveForm.rtfText.RightMargin = lngPaperWidth
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
cboPaper.ListIndex = 0
cboUnits.ListIndex = 0
End Sub

Private Sub optPaper_Click(Index As Integer)
If Index = 0 Then
    txtWidth.Enabled = False
    txtWidth.ForeColor = &HC0C0C0
    cboPaper.Enabled = True
    cboUnits.Enabled = False
    smcScaleMode = vbInches
    cboPaper_Click
    chkLandscape.Enabled = True
Else
    txtWidth.Enabled = True
    txtWidth.ForeColor = &H0
    cboPaper.Enabled = False
    cboUnits.Enabled = True
    cboUnits_Click
    txtWidth_Change
    chkLandscape.Enabled = False
End If
End Sub

Private Sub txtWidth_Change()
On Error Resume Next
lngPaperWidth = CInt(txtWidth.text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("0") To Asc("9")
Case Asc(".")
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
