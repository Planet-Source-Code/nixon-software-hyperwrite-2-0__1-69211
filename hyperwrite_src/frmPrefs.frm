VERSION 5.00
Begin VB.Form frmPrefs 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferences"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6240
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGetFonts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Get list of fonts used in document"
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
      Left            =   1208
      TabIndex        =   12
      Top             =   1680
      Width           =   4740
   End
   Begin VB.CommandButton cmdDone 
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
      Left            =   4950
      TabIndex        =   11
      Top             =   3270
      Width           =   1065
   End
   Begin VB.CheckBox chkSaveWorkspace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save/load workspace"
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
      Left            =   1208
      TabIndex        =   7
      Top             =   210
      Width           =   4770
   End
   Begin VB.CheckBox chkSymbolMatic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use SymbolMatic by default"
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
      Left            =   1208
      TabIndex        =   6
      Top             =   870
      Width           =   4755
   End
   Begin VB.CheckBox chkRecentFiles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keep List of Recent Files"
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
      Left            =   1208
      TabIndex        =   5
      Top             =   2550
      Width           =   4635
   End
   Begin VB.CheckBox chkRplcStrghtQts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto replace straight quotes when SymbolMatic is enabled"
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
      Left            =   1208
      TabIndex        =   4
      Top             =   1140
      Width           =   4755
   End
   Begin VB.CheckBox chkWarn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warn before saving as text format"
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
      Left            =   1208
      TabIndex        =   3
      Top             =   2820
      Width           =   4635
   End
   Begin VB.CheckBox chkImportPictures 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Import Pictures"
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
      Left            =   1208
      TabIndex        =   2
      Top             =   2295
      Width           =   4635
   End
   Begin VB.CheckBox chkFindStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use status bar instead of dialog after finding/replacing"
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
      Left            =   1208
      TabIndex        =   1
      Top             =   1410
      Width           =   4740
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   3270
      Width           =   1065
   End
   Begin VB.Label lblEditing 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   248
      TabIndex        =   10
      Top             =   870
      Width           =   825
   End
   Begin VB.Label lblFiles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Files:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   9
      Top             =   2295
      Width           =   825
   End
   Begin VB.Label lblGeneral 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "General:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   263
      TabIndex        =   8
      Top             =   210
      Width           =   825
   End
   Begin VB.Line lnLine 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   240
      X2              =   6000
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line lnLine 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   233
      X2              =   5993
      Y1              =   660
      Y2              =   660
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdReset_Click()
Select Case CustomBox("Are you sure you want to reset the preferences file and list of recent files?", _
    "You cannot undo this operation.", _
    vbExclamation, "", "&Cancel", "&Reset", 2)
    Case 1
        ResetPrefs
        Form_Load
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
    chkSaveWorkspace.Value = DoPrefs(0, "SaveWorkspace")
    chkSymbolMatic.Value = DoPrefs(0, "DefSymbolMatic")
    chkRplcStrghtQts.Value = DoPrefs(0, "AutoReplaceStraightQuotes")
    chkImportPictures.Value = DoPrefs(0, "ImportPictures")
    chkRecentFiles.Value = DoPrefs(0, "RecentFiles")
    chkWarn.Value = DoPrefs(0, "WarnTextFormat")
    chkFindStatus.Value = DoPrefs(0, "StatusBarFind")
    chkGetFonts.Value = DoPrefs(0, "ParseFontTable")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    DoPrefs 1, "SaveWorkspace", CStr(chkSaveWorkspace.Value)
    DoPrefs 1, "DefSymbolMatic", CStr(chkSymbolMatic.Value)
    DoPrefs 1, "AutoReplaceStraightQuotes", CStr(chkRplcStrghtQts.Value)
    DoPrefs 1, "ImportPictures", CStr(chkImportPictures.Value)
    DoPrefs 1, "RecentFiles", CStr(chkRecentFiles.Value)
    DoPrefs 1, "WarnTextFormat", CStr(chkWarn.Value)
    DoPrefs 1, "StatusBarFind", CStr(chkFindStatus.Value)
    DoPrefs 1, "ParseFontTable", CStr(chkGetFonts.Value)
End Sub

