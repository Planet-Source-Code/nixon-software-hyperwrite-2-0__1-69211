VERSION 5.00
Begin VB.Form frmCharacters 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Character"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   1695
   Icon            =   "frmCharacters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   4350
      Left            =   90
      TabIndex        =   3
      Top             =   150
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Note"
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtCode 
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
      HideSelection   =   0   'False
      Left            =   630
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "1"
      Top             =   4530
      Width           =   555
   End
   Begin VB.Label lblCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   465
   End
End
Attribute VB_Name = "frmCharacters"
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
Private Sub cmdCodeOK_Click()
On Error GoTo 10
If txtCode.Text = "0" Or CInt(txtCode.Text) > 255 Then
CustomBox "Invalid Number", "Character codes must be between 1 and 255.", vbExclamation, "", "", "OK"
txtCode.SelStart = 0
txtCode.SelLength = Len(txtCode.Text)
txtCode.SetFocus
End If
fMainForm.ActiveForm.rtfText.SelText = Chr(CInt(txtCode.Text))
10:
End Sub

Private Sub Command1_Click()
CustomBox "Note:", "To insert characters within a text box, hold ALT while pressing 0 plus the code for the character on the number pad, then release ALT. Numbers above 255 are Unicode.", vbInformation, "", "", "&OK"
End Sub

Private Sub Form_Load()
On Error GoTo 5
Dim i As Integer
For i = 1 To 255
If i = 9 Then
lstList.AddItem Chr(i) & i
Else
lstList.AddItem Chr(i) & Chr(9) & i
End If
Next
5:
End Sub

Private Sub lstList_Click()
On Error GoTo 5
txtCode.Text = lstList.ListIndex + 1
5:
End Sub

Private Sub lstList_DblClick()
On Error Resume Next
fMainForm.ActiveForm.rtfText.SelText = Chr(lstList.ListIndex + 1)
End Sub

Private Sub txtCode_Change()
If Not IsNumeric(txtCode.Text) Then
    txtCode.Text = 0
    txtCode.SelStart = 1
Else
    If CInt(txtCode.Text) > 255 Then txtCode.Text = "255"
End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then fMainForm.ActiveForm.rtfText.SelText = Chr(txtCode.Text)
End Sub
