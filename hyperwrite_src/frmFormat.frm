VERSION 5.00
Begin VB.Form frmFormat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Paragraph"
   ClientHeight    =   2325
   ClientLeft      =   2760
   ClientTop       =   3690
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLSp 
      Caption         =   ">"
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
      Left            =   2790
      TabIndex        =   10
      Top             =   555
      Width           =   345
   End
   Begin VB.CommandButton cmdLSm 
      Caption         =   "<"
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
      Left            =   1620
      TabIndex        =   9
      Top             =   555
      Width           =   345
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
      Height          =   360
      Left            =   2310
      TabIndex        =   8
      Top             =   1770
      Width           =   840
   End
   Begin VB.OptionButton optSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Inches"
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
      Index           =   0
      Left            =   255
      TabIndex        =   7
      Top             =   195
      Value           =   -1  'True
      Width           =   990
   End
   Begin VB.OptionButton optSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Centimetres"
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
      Index           =   1
      Left            =   1425
      TabIndex        =   6
      Top             =   195
      Width           =   1245
   End
   Begin VB.TextBox txtRightIndent 
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
      Height          =   300
      Left            =   1635
      TabIndex        =   5
      Top             =   1275
      Width           =   1515
   End
   Begin VB.TextBox txtLeftIndent 
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
      Height          =   300
      Left            =   1635
      TabIndex        =   4
      Top             =   915
      Width           =   1515
   End
   Begin VB.TextBox txtLineSpacing 
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
      Height          =   300
      Left            =   1980
      TabIndex        =   3
      Top             =   555
      Width           =   795
   End
   Begin VB.Label lblLnSpacing 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Line Spacing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label lblRightIndent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Right Indent:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label lblLeftIndent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left Indent:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   255
      TabIndex        =   0
      Top             =   960
      Width           =   1380
   End
End
Attribute VB_Name = "frmFormat"
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

Private Sub cmdLSm_Click()
If CSng(txtLineSpacing.Text) <= -1 Then
txtLineSpacing.Text = "-1"
Exit Sub
End If
txtLineSpacing.Text = (CSng(txtLineSpacing.Text) - 0.5)
End Sub

Private Sub cmdLSp_Click()
txtLineSpacing.Text = (CSng(txtLineSpacing.Text) + 0.5)
End Sub

Private Sub cmdOK_Click()
SetFormat
Unload Me
End Sub
Private Sub SetFormat()
On Error Resume Next
Dim intLineSpacing As Integer
Dim lngLeftIndent As Long
Dim lngRightIndent As Long
Dim lngOffset As Long
'Dim FirstLine As String
intLineSpacing = CInt(txtLineSpacing.Text * 240)
lngLeftIndent = CLng(txtLeftIndent.Text) * 1440
lngRightIndent = CLng(txtRightIndent.Text) * 1440
'FirstLine = "fi" & CLng(txtFirstLineText1.Text) * 20
ChangeLineSpacing intLineSpacing
fMainForm.ActiveForm.rtfText.SelIndent = lngLeftIndent
fMainForm.ActiveForm.rtfText.SelRightIndent = lngRightIndent
'fMainForm.ActiveForm.rtfText.TextRTF = Replace(fMainForm.ActiveForm.rtfText.TextRTF, "\pard", "\pard\" & FirstLine)
fMainForm.ActiveForm.rtfText.SelCharOffset = lngOffset
End Sub

Private Sub Form_Load()
On Error Resume Next
txtLineSpacing.Text = fMainForm.ActiveForm.rtfText.SelCharOffset / 240
txtLeftIndent.Text = fMainForm.ActiveForm.rtfText.SelIndent / 1440
txtRightIndent.Text = fMainForm.ActiveForm.rtfText.SelRightIndent / 1440
End Sub

Private Sub optSystem_Click(Index As Integer)
If Index = 1 Then
 optSystem(0).Value = False
 optSystem(1).Value = True
 txtRightIndent.Text = CLng(txtRightIndent.Text) * 2.5
 txtLeftIndent.Text = CLng(txtLeftIndent.Text) * 2.5
Else
 optSystem(0).Value = True
 optSystem(1).Value = False
 txtRightIndent.Text = CLng(txtRightIndent.Text) * 0.4
 txtLeftIndent.Text = CLng(txtLeftIndent.Text) * 0.4
End If
End Sub


Private Sub ChangeLineSpacing(intLineSpacing As Integer)
Dim intSLPos As Integer
With fMainForm.ActiveForm.rtfText
    If .SelText <> "" Then
        intSLPos = InStr(1, .SelRTF, "\sl")
    Else
        intSLPos = InStr(1, .TextRTF, "\sl")
    End If
    If intSLPos <> 0 Then
        Dim strTemp As String
        If .SelText <> "" Then
            strTemp = .SelRTF
        Else
            strTemp = .TextRTF
        End If
        strTemp = Replace(strTemp, "\slmult", "")
        strTemp = Replace(strTemp, "\sl", "\sl" & intLineSpacing & "\slmult1")
        If .SelText <> "" Then
            .SelRTF = strTemp
        Else
            .TextRTF = strTemp
        End If
    Else
        If .SelText <> "" Then
            .SelRTF = Replace(.SelRTF, "\pard", "\pard\sl" & intLineSpacing & "\slmult0")
        Else
            .TextRTF = Replace(.TextRTF, "\pard", "\pard\sl" & intLineSpacing & "\slmult0")
    
        End If
    End If
End With
End Sub

Function ReplaceWord(strText As String, _
strFind As String, _
strReplace As String) As String

' This function searches a string for a word and replaces it.
' You can use a wildcard mask to specify the search string.

Dim astrText() As String
Dim lngCount As Long

' Split the string at specified delimiter.
astrText = Split(strText)

' Loop through array, performing comparison
' against wildcard mask.
For lngCount = LBound(astrText) To UBound(astrText)
If astrText(lngCount) Like strFind Then
' If array element satisfies wildcard search,
' replace it.
astrText(lngCount) = strReplace
End If
Next
' Join string, using same delimiter.
ReplaceWord = Join(astrText)
End Function

Private Sub txtLineSpacing_Change()
If Not IsNumeric(txtLineSpacing.Text) Then
    txtLineSpacing.Text = "0"
    txtLineSpacing.SelStart = 1
End If
End Sub
