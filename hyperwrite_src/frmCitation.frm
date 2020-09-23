VERSION 5.00
Begin VB.Form frmCitation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Citation"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4905
   Icon            =   "frmCitation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
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
      Left            =   270
      TabIndex        =   11
      Top             =   3780
      Width           =   1095
   End
   Begin VB.OptionButton optSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encyclopedia"
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
      Index           =   2
      Left            =   2220
      TabIndex        =   10
      Top             =   3390
      Width           =   1335
   End
   Begin VB.OptionButton optSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Website"
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
      Index           =   1
      Left            =   1140
      TabIndex        =   9
      Top             =   3390
      Width           =   975
   End
   Begin VB.OptionButton optSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Book"
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
      Index           =   0
      Left            =   270
      TabIndex        =   8
      Top             =   3390
      Width           =   765
   End
   Begin VB.TextBox txtArticle 
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
      Left            =   1125
      TabIndex        =   6
      Top             =   2310
      Width           =   3510
   End
   Begin VB.TextBox txtWebsite 
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
      Left            =   1125
      TabIndex        =   5
      Top             =   1950
      Width           =   3510
   End
   Begin VB.TextBox txtDate 
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
      Left            =   1125
      TabIndex        =   3
      Top             =   1230
      Width           =   3510
   End
   Begin VB.TextBox txtPublisher 
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
      Left            =   1125
      TabIndex        =   4
      Top             =   1590
      Width           =   3510
   End
   Begin VB.TextBox txtPlace 
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
      Left            =   1125
      TabIndex        =   7
      Top             =   2670
      Width           =   3510
   End
   Begin VB.TextBox txtAuthor 
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
      Left            =   1125
      TabIndex        =   1
      Top             =   510
      Width           =   3510
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   1125
      TabIndex        =   2
      Top             =   870
      Width           =   3510
   End
   Begin VB.Label lblCitationSrc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Citation Source:"
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
      Left            =   270
      TabIndex        =   19
      Top             =   3120
      Width           =   3285
   End
   Begin VB.Label lblArticle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Article:"
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
      Left            =   255
      TabIndex        =   13
      Top             =   2325
      Width           =   750
   End
   Begin VB.Label lblStyle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MLA Style"
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
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   4365
   End
   Begin VB.Label lblWeb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Website:"
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
      Left            =   255
      TabIndex        =   12
      Top             =   1965
      Width           =   750
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date:"
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
      Left            =   255
      TabIndex        =   14
      Top             =   1245
      Width           =   750
   End
   Begin VB.Label lblPublisher 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Publisher:"
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
      Left            =   255
      TabIndex        =   15
      Top             =   1605
      Width           =   750
   End
   Begin VB.Label lblPlace 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Place:"
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
      Left            =   255
      TabIndex        =   16
      Top             =   2685
      Width           =   750
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Author:"
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
      Left            =   255
      TabIndex        =   18
      Top             =   525
      Width           =   750
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title:"
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
      Left            =   255
      TabIndex        =   17
      Top             =   885
      Width           =   750
   End
End
Attribute VB_Name = "frmCitation"
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
Dim finaltext(2) As String

Private Sub cmdInsert_Click()
If optSource.Item(0).Value = True Then
    finaltext(0) = txtAuthor.Text & ". "
    finaltext(1) = txtTitle.Text & "."
    finaltext(2) = " " & txtPlace.Text & ": " & txtPublisher.Text & ", " & txtDate.Text & "."
End If
If optSource.Item(1).Value = True Then
    finaltext(0) = txtAuthor.Text & ". " & Chr(147) & txtArticle.Text & "." & Chr(148) & " "
    finaltext(1) = txtTitle.Text
    finaltext(2) = ". " & txtDate.Text & ". <" & txtWebsite.Text & ">."
End If
If optSource.Item(2).Value = True Then
    finaltext(0) = txtAuthor.Text & ". " & Chr(147) & txtArticle.Text & "." & Chr(148) & " "
    finaltext(1) = txtTitle.Text
    finaltext(2) = ". " & txtDate.Text & "."
End If
WriteCitations
Unload Me
End Sub

Private Sub Form_Load()
optSource.Item(0).Value = True
End Sub

Private Sub optSource_Click(Index As Integer)
Select Case Index
Case 0
    txtAuthor.Visible = True
    txtTitle.Visible = True
    txtPlace.Visible = True
    txtPublisher.Visible = True
    txtDate.Visible = True
    txtWebsite.Visible = False
    txtArticle.Visible = False
    txtPlace.Move txtPublisher.Left, txtPublisher.Top + txtPublisher.Height + 45
    lblPlace.Visible = True
    lblPlace.Move lblPublisher.Left, lblPublisher.Top + lblPublisher.Height + 30
    lblArticle.Visible = False
    Me.lblWeb.Visible = False
    lblPublisher.Visible = True
    lblCitationSrc.Move lblPlace.Left, lblPlace.Top + lblPlace.Height + 45
    MoveBottom
    Me.Height = cmdInsert.Top + cmdInsert.Height + 800
    Rem    author
    Rem    Title
    Rem    Date
    Rem    publisher
    Rem    place
    txtAuthor.TabIndex = 1
    txtTitle.TabIndex = 2
    txtDate.TabIndex = 3
    txtPublisher.TabIndex = 4
    txtPlace.TabIndex = 5
    optSource(0).TabIndex = 6
    optSource(1).TabIndex = 7
    optSource(2).TabIndex = 8
    cmdInsert.TabIndex = 9
Case 1
    txtAuthor.Visible = True
    txtTitle.Visible = True
    txtPlace.Visible = False
    txtPublisher.Visible = True
    txtDate.Visible = True
    txtWebsite.Visible = True
    txtArticle.Visible = True
    lblPlace.Visible = False
    lblWeb.Visible = True
    lblArticle.Visible = True
    txtWebsite.Move txtDate.Left, txtDate.Top + txtDate.Height + 45
    lblWeb.Move lblDate.Left, lblDate.Top + lblDate.Height + 30
    txtArticle.Move txtWebsite.Left, txtWebsite.Top + txtWebsite.Height + 45
    lblArticle.Move lblWeb.Left, lblWeb.Top + lblWeb.Height + 30
    lblCitationSrc.Move lblArticle.Left, lblArticle.Top + lblArticle.Height + 45
    MoveBottom
    Me.Height = cmdInsert.Top + cmdInsert.Height + 800
    txtAuthor.TabIndex = 1
    txtTitle.TabIndex = 2
    txtDate.TabIndex = 3
    txtWebsite.TabIndex = 4
    txtArticle.TabIndex = 5
    optSource(0).TabIndex = 6
    optSource(1).TabIndex = 7
    optSource(2).TabIndex = 8
    cmdInsert.TabIndex = 9
Case 2
    txtAuthor.Visible = True
    txtTitle.Visible = True
    txtPlace.Visible = False
    txtPublisher.Visible = False
    txtDate.Visible = True
    txtWebsite.Visible = False
    txtArticle.Visible = True
    lblPublisher.Visible = False
    lblWeb.Visible = False
    lblPlace.Visible = False
    lblPlace.Move lblDate.Left, lblDate.Top + lblDate.Height + 30
    lblArticle.Move lblDate.Left, lblDate.Top + lblDate.Height + 30
    txtArticle.Move txtDate.Left, txtDate.Top + txtDate.Height + 45
    lblArticle.Visible = True
    lblCitationSrc.Move lblArticle.Left, lblArticle.Top + lblArticle.Height + 45
    MoveBottom
    Me.Height = cmdInsert.Top + cmdInsert.Height + 800
    txtAuthor.TabIndex = 1
    txtTitle.TabIndex = 2
    txtDate.TabIndex = 3
    txtArticle.TabIndex = 4
    optSource(0).TabIndex = 5
    optSource(1).TabIndex = 6
    optSource(2).TabIndex = 7
    cmdInsert.TabIndex = 8
End Select
End Sub
Private Sub WriteCitations()
If fMainForm.ActiveForm Is Nothing Then Exit Sub
With fMainForm.ActiveForm.rtfText
.SelText = finaltext(0)
.SelUnderline = True
.SelText = finaltext(1)
.SelUnderline = False
.SelText = finaltext(2)
End With
End Sub

Private Sub MoveBottom()
    optSource.Item(0).Move lblCitationSrc.Left, lblCitationSrc.Top + lblCitationSrc.Height + 45
    optSource.Item(1).Move optSource.Item(0).Left + optSource.Item(0).Width + 15, optSource(0).Top
    optSource.Item(2).Move optSource.Item(1).Left + optSource.Item(1).Width + 15, optSource(0).Top
    cmdInsert.Top = optSource.Item(0).Top + optSource.Item(0).Height + 45
End Sub
