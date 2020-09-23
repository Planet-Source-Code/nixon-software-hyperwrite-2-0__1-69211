VERSION 5.00
Begin VB.Form frmCopylist 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copy List"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6930
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6045
      TabIndex        =   6
      Top             =   105
      Width           =   825
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   900
      TabIndex        =   5
      Top             =   120
      Width           =   5070
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
      Height          =   330
      Left            =   90
      TabIndex        =   3
      Top             =   4605
      Width           =   975
   End
   Begin VB.CommandButton bnClear 
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
      Height          =   330
      Left            =   4845
      TabIndex        =   2
      Top             =   4605
      Width           =   975
   End
   Begin VB.CommandButton bnClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5865
      TabIndex        =   1
      Top             =   4605
      Width           =   975
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
      Height          =   3570
      Left            =   15
      TabIndex        =   0
      Top             =   540
      Width           =   6870
   End
   Begin VB.Label lblLengthSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Length: 0 / Size: 0 bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   945
      TabIndex        =   8
      Top             =   4170
      Width           =   5940
   End
   Begin VB.Label lblDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Details:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   7
      Top             =   4170
      Width           =   765
   End
   Begin VB.Label lblAddItem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Item:"
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
      Left            =   45
      TabIndex        =   4
      Top             =   135
      Width           =   855
   End
End
Attribute VB_Name = "frmCopylist"
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
Private Sub bnClear_Click()
lstList.Clear
End Sub

Private Sub bnClose_Click()
frmCopylist.Hide
End Sub

Private Sub cmdAdd_Click()
lstList.AddItem txtItem.Text
End Sub

Private Sub cmdRemove_Click()
lstList.RemoveItem (lstList.ListIndex)
End Sub

Private Sub lstList_Click()
lblLengthSize.Caption = "Length: " & Len(lstList.List(lstList.ListIndex)) & " / Size: " & LenB(lstList.List(lstList.ListIndex)) & " bytes"
End Sub

Private Sub lstList_DblClick()
fMainForm.ActiveForm.rtfText.SelText = lstList.List(lstList.ListIndex)
If bLiveWC = True Then fMainForm.StatusBar.Panels(1).Text = WordCount(fMainForm.ActiveForm.rtfText.Text) & " words"
frmCopylist.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd_Click
End Sub
