VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReadMode 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Read Mode"
   ClientHeight    =   3090
   ClientLeft      =   90
   ClientTop       =   255
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "frmReadMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfReadMode 
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1349
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmReadMode.frx":038A
   End
End
Attribute VB_Name = "frmReadMode"
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
Private Sub Form_Load()
On Error GoTo 10
Me.Width = Screen.Width
Me.Height = Screen.Height
rtfReadMode.TextRTF = fMainForm.ActiveForm.rtfText.TextRTF
rtfReadMode.SelStart = 0
rtfReadMode.SelLength = Len(rtfReadMode.Text)
rtfReadMode.SelIndent = fMainForm.ActiveForm.rtfText.SelIndent
rtfReadMode.SelRightIndent = fMainForm.ActiveForm.rtfText.SelRightIndent
rtfReadMode.BackColor = fMainForm.ActiveForm.rtfText.BackColor
rtfReadMode.SelProtected = True
rtfReadMode.SelStart = Len(rtfReadMode.Text)
rtfReadMode.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
rtfReadMode.Locked = True
10:
End Sub
