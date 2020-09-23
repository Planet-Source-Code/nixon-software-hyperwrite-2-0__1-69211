VERSION 5.00
Begin VB.Form frmReplaceFonts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Replace Fonts"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   3015
      TabIndex        =   5
      Top             =   1125
      Width           =   1020
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
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
      Height          =   330
      Left            =   1890
      TabIndex        =   4
      Top             =   1125
      Width           =   1020
   End
   Begin VB.ComboBox cboFonts 
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
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   615
      Width           =   2955
   End
   Begin VB.ComboBox cboCurrFonts 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   195
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "With:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   255
      TabIndex        =   1
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Replace:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "frmReplaceFonts"
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
        ' your About dialog.                                     '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdReplace_Click()
Dim intFontTableStart As Integer, strBefore As String
    With fMainForm.ActiveForm.rtfText
        intFontTableStart = InStr(1, .TextRTF, "{\fonttbl")
        strBefore = Left(.TextRTF, intFontTableStart - 1)
        'intFontTableEnd = InStr(intFontTableStart, .TextRTF, "}}")
        .TextRTF = strBefore & Replace(.TextRTF, cboCurrFonts.List(cboCurrFonts.ListIndex) & ";", cboFonts.Text & ";", intFontTableStart, 1)
    End With
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To Screen.FontCount - 1
    cboFonts.AddItem Screen.Fonts(i)
Next
For i = 0 To GetLastFontNum
    cboCurrFonts.AddItem ParseFontTable(i)
Next
cboFonts.ListIndex = 0
cboCurrFonts.ListIndex = 0
End Sub

