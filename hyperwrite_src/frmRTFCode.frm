VERSION 5.00
Begin VB.Form frmRTFCode 
   Caption         =   "RTF Code"
   ClientHeight    =   3840
   ClientLeft      =   90
   ClientTop       =   315
   ClientWidth     =   8280
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmRTFCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRTF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8265
   End
End
Attribute VB_Name = "frmRTFCode"
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
Dim bChanged As Boolean

Private Sub Form_Load()
    If fMainForm.mnuViewRTF.Tag = "Whole" Then
        frmRTFCode.txtRTF.Text = fMainForm.ActiveForm.rtfText.TextRTF
    Else
        frmRTFCode.txtRTF.Text = fMainForm.ActiveForm.rtfText.SelRTF
    End If
    bChanged = False
End Sub

Private Sub Form_Resize()
On Error GoTo 10
    On Error Resume Next
    txtRTF.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
10:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim intMsgReturn As Integer
If bChanged = True Then
    intMsgReturn = CustomBox("You have made changes to the RTF code in this window. Do you want to apply the code in your document?", "If you don" & sApostrophe & "t apply the code, the changes you made to it won" & sApostrophe & "t affect the document.", vbExclamation, "Don" & sApostrophe & "t Apply", "Cancel", "Apply")
    If intMsgReturn = 1 Then
        If txtRTF.Tag = "Whole" Then
            fMainForm.ActiveForm.rtfText.TextRTF = txtRTF.Text
        Else
            fMainForm.ActiveForm.rtfText.SelRTF = txtRTF.Text
        End If
    End If
    If intMsgReturn = 2 Then Cancel = 1
End If
End Sub

Private Sub txtRTF_Change()
bChanged = True
End Sub
