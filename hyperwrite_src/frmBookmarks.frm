VERSION 5.00
Begin VB.Form frmBookmarks 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bookmarks"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3030
   Icon            =   "frmBookmarks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bnClear 
      Caption         =   "&Clear "
      Height          =   345
      Left            =   173
      TabIndex        =   2
      Top             =   2947
      Width           =   735
   End
   Begin VB.CommandButton bnClose 
      Caption         =   "C&lose"
      Default         =   -1  'True
      Height          =   345
      Left            =   2123
      TabIndex        =   1
      Top             =   2947
      Width           =   735
   End
   Begin VB.ListBox lstList 
      Height          =   2790
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   2685
   End
End
Attribute VB_Name = "frmBookmarks"
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
Unload Me
End Sub


Private Sub lstList_Click()
On Error GoTo 10
fMainForm.ActiveForm.rtfText.SelStart = lstList.List(lstList.ListIndex)
frmBookmarks.Hide
10:
End Sub

Private Sub lstList_DblClick()
On Error GoTo 10
fMainForm.ActiveForm.rtfText.SelStart = lstList.List(lstList.ListIndex)
frmBookmarks.Hide
10:
End Sub
