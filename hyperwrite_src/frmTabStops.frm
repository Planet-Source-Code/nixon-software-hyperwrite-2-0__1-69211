VERSION 5.00
Begin VB.Form frmTabStops 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tab Stops"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3765
   Icon            =   "frmTabStops.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   1455
      TabIndex        =   5
      Top             =   2295
      Width           =   915
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
      Height          =   345
      Left            =   435
      TabIndex        =   4
      Top             =   2295
      Width           =   915
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
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
      Left            =   2655
      TabIndex        =   3
      Top             =   225
      Width           =   645
   End
   Begin VB.TextBox txtSet 
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
      Left            =   435
      MaxLength       =   5
      TabIndex        =   2
      Top             =   240
      Width           =   1305
   End
   Begin VB.ListBox lstList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmTabStops.frx":5F32
      Left            =   435
      List            =   "frmTabStops.frx":5F34
      TabIndex        =   1
      Top             =   615
      Width           =   2865
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   2475
      TabIndex        =   0
      Top             =   2295
      Width           =   855
   End
   Begin VB.Label lblInches 
      BackStyle       =   0  'Transparent
      Caption         =   "inches"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1830
      TabIndex        =   6
      Top             =   300
      Width           =   645
   End
End
Attribute VB_Name = "frmTabStops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
lstList.Clear
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
fMainForm.ActiveForm.rtfText.SelTabCount = lstList.ListCount
For i = 0 To lstList.ListCount - 1
    fMainForm.ActiveForm.rtfText.SelTabs(i) = (lstList.List(i) * 1440)
Next
Unload Me
End Sub

Private Sub cmdRemove_Click()
lstList.RemoveItem (lstList.ListIndex)
End Sub

Private Sub cmdSet_Click()
Dim i As Integer
For i = 0 To lstList.ListCount
    If txtSet.Text = lstList.List(i) Then Exit Sub
Next
lstList.AddItem txtSet.Text
End Sub

Private Sub Form_Load()
Dim i As Integer, intStartPos As Integer, intEndPos As Integer
intEndPos = 1
For i = 0 To FindOccurrences(fMainForm.ActiveForm.rtfText.TextRTF, "\tx")
    intStartPos = InStr(intEndPos, fMainForm.ActiveForm.rtfText.TextRTF, "\tx")
    intEndPos = InStr(intStartPos + 2, fMainForm.ActiveForm.rtfText.TextRTF, "\")
    lstList.AddItem Format((Mid(fMainForm.ActiveForm.rtfText.TextRTF, intStartPos + 3, intEndPos - intStartPos - 3)) / 1440, "#.00")
Next
End Sub

Private Sub txtSet_Change()
On Error GoTo 10
Dim i As Integer
i = CInt(txtSet.Text)
Exit Sub
10:
If txtSet.Text = "" Then Exit Sub
txtSet.Text = "0"
txtSet.SelStart = Len(txtSet.Text)
End Sub

Private Sub txtSet_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57
    Case vbKeyBack
    Case 46
    Case Else
        KeyAscii = 0
End Select
End Sub
