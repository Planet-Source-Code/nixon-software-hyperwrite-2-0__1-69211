VERSION 5.00
Begin VB.Form frmTables 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Table…"
   ClientHeight    =   2670
   ClientLeft      =   2760
   ClientTop       =   3735
   ClientWidth     =   2670
   Icon            =   "frmTables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRows 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   9
      Text            =   "1"
      Top             =   1680
      Width           =   765
   End
   Begin VB.ComboBox cboCell 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTables.frx":5F32
      Left            =   1740
      List            =   "frmTables.frx":5F39
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   945
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Done"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1650
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Top             =   540
      Width           =   765
   End
   Begin VB.TextBox txtWidth 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Text            =   "1"
      Top             =   1320
      Width           =   765
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A&dd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rows:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1710
      Width           =   1395
   End
   Begin VB.Label lblSelCell 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected Cell:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblCell 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cells: 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label lblWidth 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cell Width (inches):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   1350
      Width           =   1485
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
        ' Hyperwrite from NIXON                                  '
        ' Copyright (C) 2004-2007 NIXON Software Corporation.    '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
        ' You may use this code freely in your own applications. '
        ' If you are distributing your code/application(s), it   '
        ' would be greatly appreciated if you credit NIXON in    '
        ' your About dialog. Please note that portions of this   '
        ' code belongs to other people. For more details, please '
        ' view the About dialog.                                 '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit
Dim dblCellWidth(39) As Double
Dim btCells As Byte

Private Sub cmdRemove_Click()
If btCells > 1 Then
    btCells = btCells - 1
    cboCell.RemoveItem (cboCell.ListIndex)
    lblCell.Caption = "Cells: " & btCells
End If
End Sub

Private Sub cbocell_Click()
txtWidth.Text = dblCellWidth(cboCell.ListIndex)
End Sub

Private Sub Form_Load()
btCells = 1
cboCell.ListIndex = 0
End Sub

Private Sub cmdAdd_Click()
If cboCell.ListCount < 40 Then
    btCells = btCells + 1
    cboCell.AddItem btCells
    dblCellWidth(cboCell.ListIndex) = 1
    lblCell.Caption = "Cells: " & btCells
Else
    CustomBox "Could not add cell because there are too many cells specified.", "The maximum amount of cells is 40.", vbExclamation, "", "", "OK"
End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo 10
Dim sRTF(5) As String
Dim sTemp(3) As String
Dim i As Integer
fMainForm.StatusBar.Style = sbrSimple
fMainForm.StatusBar.SimpleText = "Creating Table..."
fMainForm.ActiveForm.ScaleMode = vbPixels
For i = 1 To UBound(dblCellWidth)
    dblCellWidth(i) = dblCellWidth(i - 1) + dblCellWidth(i)
Next
sRTF(0) = "\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw10 "
sRTF(1) = "\trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 "
sRTF(2) = "\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx" & CLng(dblCellWidth(0) * 1420)
    If btCells > 1 Then
        For i = 1 To btCells - 1
        sRTF(3) = sRTF(3) & "\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx" & CLng(dblCellWidth(i) * 1420)
        sTemp(0) = sTemp(0) & "\cell"
        Next
    End If
For i = 0 To CInt(txtRows.Text) - 1
    sTemp(1) = sTemp(1) & "\row"
Next
sRTF(4) = "\pard\intbl\f0\fs24" & "\cell" & sTemp(0) & sTemp(1) & vbNewLine & "\pard }"
For i = 0 To 5
    sRTF(5) = sRTF(5) + vbNewLine + sRTF(i)
Next
fMainForm.ActiveForm.rtfText.SelText = "\tr" & fMainForm.ActiveForm.rtfText.SelText
fMainForm.ActiveForm.rtfText.TextRTF = Replace(fMainForm.ActiveForm.rtfText.TextRTF, "\\tr" & fMainForm.ActiveForm.rtfText.SelText, sRTF(5))
fMainForm.StatusBar.Style = sbrNormal
fMainForm.StatusBar.SimpleText = "Ready"
Unload Me
10:
End Sub

Private Sub txtRows_Change()
If Not IsNumeric(txtRows.Text) Then
    txtRows.Text = "0"
    txtRows.SelStart = 1
Else
    If CDbl(txtRows.Text) > 32767 Then txtRows.Text = 32767
End If
End Sub

Private Sub txtRows_KeyDown(KeyCode As Integer, Shift As Integer)
Shift = 0
Select Case KeyCode
    Case 48 To 57 Or 108
    Case Else
        KeyCode = 0
End Select
End Sub

Private Sub txtWidth_Change()
If Not IsNumeric(txtWidth.Text) Then
    dblCellWidth(cboCell.List(cboCell.ListIndex) - 1) = 0
    txtWidth.Text = "0"
    txtWidth.SelStart = 1
    Exit Sub
End If
dblCellWidth(cboCell.List(cboCell.ListIndex) - 1) = CDbl(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyDown(KeyCode As Integer, Shift As Integer)
If Not IsNumeric(txtWidth.Text) Then KeyCode = 0
End Sub

