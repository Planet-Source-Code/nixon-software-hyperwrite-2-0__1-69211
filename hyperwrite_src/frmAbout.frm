VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About NIXON Hyperwrite"
   ClientHeight    =   4950
   ClientLeft      =   3120
   ClientTop       =   2685
   ClientWidth     =   5700
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":2CFA
   ScaleHeight     =   3416.579
   ScaleMode       =   0  'User
   ScaleWidth      =   5352.597
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIXON Hyperwrite 2.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   615
      TabIndex        =   3
      Top             =   2055
      Width           =   4470
   End
   Begin VB.Image imgHW2 
      Height          =   630
      Left            =   2025
      Picture         =   "frmAbout.frx":49EC
      Top             =   675
      Width           =   3165
   End
   Begin VB.Image imgLogo 
      Height          =   1680
      Left            =   315
      Picture         =   "frmAbout.frx":B288
      Top             =   150
      Width           =   1485
   End
   Begin VB.Label lblWebsite 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit us online at members.shaw.ca/nixon.com"
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
      Height          =   270
      Left            =   645
      TabIndex        =   2
      Top             =   3660
      Width           =   4395
   End
   Begin VB.Image imgNixonLogo 
      Height          =   510
      Left            =   2100
      Picture         =   "frmAbout.frx":1360C
      Top             =   4140
      Width           =   1500
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright © 2004-2007 NIXON Software Corporation."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   630
      TabIndex        =   1
      Top             =   2625
      Width           =   4425
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   615
      TabIndex        =   0
      Top             =   2340
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
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
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = "Copyright © 2004-2007 NIXON Software Corporation." & _
        vbNewLine & "Portions © 2006 York Technologies Ltd." & vbNewLine & _
                    "Portions © 1985-1996 Microsoft Corporation." & vbNewLine & _
                    "Portions © 2000 Seagate Software, Inc."
End Sub
