VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H007E6D5C&
   Caption         =   "Hyperwrite"
   ClientHeight    =   9435
   ClientLeft      =   -45
   ClientTop       =   555
   ClientWidth     =   13470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "frmMain.frx":6872
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbSymbols 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   25
      Top             =   8505
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   39
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "×"
            Object.ToolTipText     =   "×"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "÷"
            Object.ToolTipText     =   "÷"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "½"
            Description     =   "½"
            Object.ToolTipText     =   "½"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "¼"
            Description     =   "¼"
            Object.ToolTipText     =   "¼"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "¾"
            Description     =   "¾"
            Object.ToolTipText     =   "¾"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "±"
            Description     =   "±"
            Object.ToolTipText     =   "±"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "°"
            Description     =   "°"
            Object.ToolTipText     =   "°"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "€"
            Description     =   "€"
            Object.ToolTipText     =   "€"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "£"
            Description     =   "£"
            Object.ToolTipText     =   "£"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "¥"
            Description     =   "¥"
            Object.ToolTipText     =   "¥"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "¢"
            Description     =   "¢"
            Object.ToolTipText     =   "¢"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "²"
            Description     =   "²"
            Object.ToolTipText     =   "²"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "³"
            Description     =   "³"
            Object.ToolTipText     =   "³"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "¿"
            Description     =   "¿"
            Object.ToolTipText     =   "¿"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "¡"
            Description     =   "¡"
            Object.ToolTipText     =   "¡"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "à"
            Description     =   "à"
            Object.ToolTipText     =   "à"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "â"
            Description     =   "â"
            Object.ToolTipText     =   "â"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "á"
            Description     =   "á"
            Object.ToolTipText     =   "á"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ã"
            Description     =   "ã"
            Object.ToolTipText     =   "ã"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ä"
            Description     =   "ä"
            Object.ToolTipText     =   "ä"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "æ"
            Description     =   "æ"
            Object.ToolTipText     =   "æ"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ç"
            Description     =   "ç"
            Object.ToolTipText     =   "ç"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "è"
            Description     =   "è"
            Object.ToolTipText     =   "è"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "é"
            Description     =   "é"
            Object.ToolTipText     =   "é"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ê"
            Description     =   "ê"
            Object.ToolTipText     =   "ê"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ë"
            Description     =   "ë"
            Object.ToolTipText     =   "ë"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ì"
            Description     =   "ì"
            Object.ToolTipText     =   "ì"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button28 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "í"
            Description     =   "í"
            Object.ToolTipText     =   "í"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button29 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "î"
            Description     =   "î"
            Object.ToolTipText     =   "î"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button30 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ï"
            Description     =   "ï"
            Object.ToolTipText     =   "ï"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button31 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ò"
            Description     =   "ò"
            Object.ToolTipText     =   "ò"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button32 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ó"
            Description     =   "ó"
            Object.ToolTipText     =   "ó"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button33 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ô"
            Description     =   "ô"
            Object.ToolTipText     =   "ô"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button34 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "õ"
            Description     =   "õ"
            Object.ToolTipText     =   "õ"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button35 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ö"
            Description     =   "ö"
            Object.ToolTipText     =   "ö"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button36 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ù"
            Description     =   "ù"
            Object.ToolTipText     =   "ù"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button37 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ú"
            Description     =   "ú"
            Object.ToolTipText     =   "ú"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button38 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "û"
            Description     =   "û"
            Object.ToolTipText     =   "û"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button39 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ü"
            Description     =   "ü"
            Object.ToolTipText     =   "ü"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrLiveWC 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   2280
      Top             =   6600
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   23
      Top             =   9105
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   582
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Ready"
            TextSave        =   "Ready"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Ln 0 Pos 0 Sel 0"
            TextSave        =   "Ln 0 Pos 0 Sel 0"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Line, Position, Selection Length"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   936
            TextSave        =   "NUM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Num Lock"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   979
            MinWidth        =   972
            TextSave        =   "CAPS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Caps Lock"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   979
            MinWidth        =   971
            TextSave        =   "SCRL"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Scroll Lock"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   847
            MinWidth        =   838
            TextSave        =   "INS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Insert Mode"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1770
      Top             =   6600
   End
   Begin ComCtl3.CoolBar cbCoolBar 
      Align           =   1  'Align Top
      Height          =   1470
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   2593
      BandBorders     =   0   'False
      _CBWidth        =   13470
      _CBHeight       =   1470
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar"
      MinHeight1      =   390
      Width1          =   3120
      NewRow1         =   0   'False
      Child2          =   "tbrFontFormat"
      MinHeight2      =   420
      Width2          =   12435
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Child3          =   "pcSlider"
      MinHeight3      =   600
      Width3          =   15375
      NewRow3         =   -1  'True
      Begin ComctlLib.Toolbar Toolbar 
         Height          =   390
         Left            =   165
         TabIndex        =   19
         Top             =   30
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   15
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "New"
               Description     =   "New"
               Object.ToolTipText     =   "New"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Open"
               Description     =   "Open"
               Object.ToolTipText     =   "Open"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Save"
               Description     =   "Save"
               Object.ToolTipText     =   "Save"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Print"
               Description     =   "Print"
               Object.ToolTipText     =   "Print"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Find/Change"
               Description     =   "Find/Change"
               Object.ToolTipText     =   "Find/Change"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Cut"
               Description     =   "Cut"
               Object.ToolTipText     =   "Cut"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Copy"
               Description     =   "Copy"
               Object.ToolTipText     =   "Copy"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Paste"
               Description     =   "Paste"
               Object.ToolTipText     =   "Paste"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Undo"
               Description     =   "Undo"
               Object.ToolTipText     =   "Undo"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Redo"
               Description     =   "Redo"
               Object.ToolTipText     =   "Redo"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Insert Image"
               Description     =   "Insert Image"
               Object.ToolTipText     =   "Insert Image"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Insert Date and Time"
               Description     =   "Insert Date and Time"
               Object.ToolTipText     =   "Insert Date and Time"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Insert Symbol"
               Description     =   "Insert Symbol"
               Object.ToolTipText     =   "Insert Symbol"
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pcSlider 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   165
         ScaleHeight     =   600
         ScaleWidth      =   13215
         TabIndex        =   12
         Top             =   840
         Width           =   13215
         Begin ComctlLib.Slider Slider 
            Height          =   315
            Left            =   -75
            TabIndex        =   21
            Top             =   105
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   556
            _Version        =   327682
            Max             =   136
         End
         Begin VB.Image imgRuler 
            Height          =   165
            Left            =   60
            Picture         =   "frmMain.frx":7574
            Top             =   420
            Width           =   12540
         End
      End
      Begin VB.PictureBox tbrFontFormat 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   165
         ScaleHeight     =   420
         ScaleWidth      =   13215
         TabIndex        =   9
         Top             =   420
         Width           =   13215
         Begin ComctlLib.Toolbar tbFormat 
            Height          =   390
            Left            =   4470
            TabIndex        =   20
            Top             =   60
            Width           =   13590
            _ExtentX        =   23971
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   15
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Bold"
                  Description     =   "Bold"
                  Object.ToolTipText     =   "Bold"
                  Object.Tag             =   ""
                  Style           =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Italic"
                  Description     =   "Italic"
                  Object.ToolTipText     =   "Italic"
                  Object.Tag             =   ""
                  Style           =   1
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Underline"
                  Description     =   "Underline"
                  Object.ToolTipText     =   "Underline"
                  Object.Tag             =   ""
                  Style           =   1
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Strikethru"
                  Description     =   "Strikethru"
                  Object.ToolTipText     =   "Strikethru"
                  Object.Tag             =   ""
                  Style           =   1
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Left"
                  Description     =   "Left"
                  Object.ToolTipText     =   "Left"
                  Object.Tag             =   ""
                  Style           =   2
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Center"
                  Description     =   "Center"
                  Object.ToolTipText     =   "Center"
                  Object.Tag             =   ""
                  Style           =   2
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Right"
                  Description     =   "Right"
                  Object.ToolTipText     =   "Right"
                  Object.Tag             =   ""
                  Style           =   2
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Bullets"
                  Description     =   "Bullets"
                  Object.ToolTipText     =   "Bullets"
                  Object.Tag             =   ""
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Superscript"
                  Description     =   "Superscript"
                  Object.ToolTipText     =   "Superscript"
                  Object.Tag             =   ""
               EndProperty
               BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Subscript"
                  Description     =   "Subscript"
                  Object.ToolTipText     =   "Subscript"
                  Object.Tag             =   ""
               EndProperty
               BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Increase Size"
                  Description     =   "Increase Size"
                  Object.ToolTipText     =   "Increase Size"
                  Object.Tag             =   ""
               EndProperty
               BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Decrease Size"
                  Description     =   "Decrease Size"
                  Object.ToolTipText     =   "Decrease Size"
                  Object.Tag             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtPreview 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cboFont 
            Height          =   315
            ItemData        =   "frmMain.frx":BD90
            Left            =   0
            List            =   "frmMain.frx":BDA3
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   90
            Width           =   555
         End
         Begin VB.ComboBox cboFontSize 
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
            ItemData        =   "frmMain.frx":BDBB
            Left            =   2850
            List            =   "frmMain.frx":BDF5
            TabIndex        =   11
            Top             =   90
            Width           =   720
         End
         Begin VB.ComboBox cboFontFace 
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
            ItemData        =   "frmMain.frx":BE41
            Left            =   570
            List            =   "frmMain.frx":BE43
            Sorted          =   -1  'True
            TabIndex        =   10
            Text            =   "Times New Roman"
            Top             =   90
            Width           =   2250
         End
         Begin VB.Label lblColor 
            BackColor       =   &H00000000&
            Height          =   285
            Left            =   3630
            TabIndex        =   24
            ToolTipText     =   "Font Color"
            Top             =   105
            Width           =   465
         End
         Begin VB.Shape shpSwatch 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   4170
            Shape           =   5  'Rounded Square
            Top             =   120
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox pctFindReplace 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      MouseIcon       =   "frmMain.frx":BE45
      Negotiate       =   -1  'True
      ScaleHeight     =   750
      ScaleWidth      =   13470
      TabIndex        =   0
      Top             =   7755
      Visible         =   0   'False
      Width           =   13470
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wrap Around"
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
         Index           =   3
         Left            =   7320
         TabIndex        =   22
         Top             =   90
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CommandButton cmdFindPrev 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prev"
         Enabled         =   0   'False
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
         Left            =   3195
         TabIndex        =   18
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton cmdSimpleReplace 
         Caption         =   "&Quick"
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
         Left            =   4740
         TabIndex        =   17
         Top             =   375
         Width           =   765
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Incremental"
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
         Index           =   2
         Left            =   5610
         TabIndex        =   15
         Top             =   330
         Width           =   1215
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Case-Sensitive"
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
         Left            =   7320
         TabIndex        =   14
         Top             =   330
         Width           =   1590
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Whole Word Only"
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
         Left            =   5610
         TabIndex        =   13
         Top             =   90
         Width           =   1590
      End
      Begin VB.CommandButton cmdReplace 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ne&xt"
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
         Left            =   3195
         TabIndex        =   7
         Top             =   375
         Width           =   765
      End
      Begin VB.CommandButton cmdFindNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         Enabled         =   0   'False
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
         Left            =   4020
         TabIndex        =   6
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdReplaceAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&All"
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
         Left            =   4020
         TabIndex        =   5
         Top             =   375
         Width           =   660
      End
      Begin VB.TextBox txtReplace 
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
         Left            =   990
         TabIndex        =   4
         Top             =   375
         Width           =   2160
      End
      Begin VB.TextBox txtFind 
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
         Height          =   285
         HideSelection   =   0   'False
         Left            =   990
         TabIndex        =   2
         Top             =   75
         Width           =   2160
      End
      Begin VB.Label lblFindReplace 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Replace w/:"
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
         Left            =   90
         TabIndex        =   3
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblFindReplace 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
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
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   795
      End
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   1140
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileNewFromClipboard 
         Caption         =   "New from &Clipboard"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenRecent 
         Caption         =   "Open Recent"
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   0
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   1
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   2
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   3
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFileOpenText 
         Caption         =   "Open as Te&xt..."
      End
      Begin VB.Menu mnuOpenBook 
         Caption         =   "Open &Book..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Close &All"
      End
      Begin VB.Menu mnuFileLine0 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Sav&e As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuFileSaveSelection 
         Caption         =   "Save Selection &To..."
      End
      Begin VB.Menu mnuFileAutoSave 
         Caption         =   "&AutoSave"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuFileLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "&Revert to Saved"
      End
      Begin VB.Menu mnuFileGetInfo 
         Caption         =   "&Get Info..."
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Insert..."
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditUndoReplace 
         Caption         =   "Undo R&eplace"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditAppend 
         Caption         =   "Appen&d"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste and Match Style"
      End
      Begin VB.Menu mnuEditPastePlain 
         Caption         =   "Paste Plain &Text"
      End
      Begin VB.Menu mnuEditCopyList 
         Caption         =   "&Add to List"
      End
      Begin VB.Menu mnuEditPasteList 
         Caption         =   "&Get from List..."
      End
      Begin VB.Menu mnuEditLine 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDelNextWord 
         Caption         =   "Delete &Next Word"
      End
      Begin VB.Menu mnuEditDelPrevWord 
         Caption         =   "Delete Previous &Word"
      End
      Begin VB.Menu mnuEditPurge 
         Caption         =   "&Purge Clipboard"
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditSelectNone 
         Caption         =   "Select &None"
      End
      Begin VB.Menu mnuEditSelUpTo 
         Caption         =   "Select &Up To Position..."
      End
      Begin VB.Menu mnuEditSelBefCur 
         Caption         =   "Select &Before Cursor"
      End
      Begin VB.Menu mnuEditSelAftCur 
         Caption         =   "Select &After Cursor"
      End
      Begin VB.Menu mnuEditLineSelect 
         Caption         =   "Select &Line"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuEditSelectNextWord 
         Caption         =   "Select &Next Word"
      End
      Begin VB.Menu mnuEditSelectPrevWord 
         Caption         =   "Select &Previous Word"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindReplace 
         Caption         =   "&Find/Replace..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditIncrementalFind 
         Caption         =   "Incremental Find"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditChgProtection 
         Caption         =   "&Change Protection"
      End
      Begin VB.Menu mnuEditTrimSpaces 
         Caption         =   "&Trim Spaces"
      End
      Begin VB.Menu mnuEditLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPreferences 
         Caption         =   "Pr&eferences..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewCoolbar 
         Caption         =   "&Coolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewFormatBar 
         Caption         =   "&Format Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewExtendedFormatting 
         Caption         =   "&Extended Formatting"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewAccentsBar 
         Caption         =   "&Symbol Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRTF 
         Caption         =   "&RTF Code"
         Begin VB.Menu mnuViewRTFCode 
            Caption         =   "for &Whole Document"
            Index           =   0
         End
         Begin VB.Menu mnuViewRTFCode 
            Caption         =   "for &Selection"
            Index           =   1
         End
      End
      Begin VB.Menu mnuViewTogglePrntNrml 
         Caption         =   "Toggle Print/Normal View"
      End
      Begin VB.Menu mnuViewPageWidth 
         Caption         =   "Page Width..."
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "&Ruler"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewToggleSlider 
         Caption         =   "&Lock Slider"
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "&Go"
      Begin VB.Menu mnuEditInsBookmark 
         Caption         =   "&Insert Bookmark"
         Shortcut        =   +^{F11}
      End
      Begin VB.Menu mnuEditGoToBookmark 
         Caption         =   "&Go To Bookmark..."
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu mnuGoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGoTo 
         Caption         =   "To &Beginning"
      End
      Begin VB.Menu mnuGoToEnd 
         Caption         =   "To &End"
      End
      Begin VB.Menu mnuGoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToLineAbove 
         Caption         =   "To Line &Above"
      End
      Begin VB.Menu mnuGoToLineBelow 
         Caption         =   "To Line B&elow"
      End
      Begin VB.Menu mnuGoToNextWord 
         Caption         =   "To &Next Word"
      End
      Begin VB.Menu mnuGoToPrevWord 
         Caption         =   "To P&revious Word"
      End
      Begin VB.Menu mnuGoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "To &Position..."
      End
      Begin VB.Menu mnuGoToLine 
         Caption         =   "To &Line#..."
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertObject 
         Caption         =   "&Image..."
      End
      Begin VB.Menu mnuInsertLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertNonbreakingSpace 
         Caption         =   "Non-breaking Space"
      End
      Begin VB.Menu mnuInsertDateandTime 
         Caption         =   "&Date and Time..."
         Begin VB.Menu mnuInsertDate 
            Caption         =   "Fixed Date"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu mnuInsertTime 
            Caption         =   "Fixed Time"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Fixed Date and Time"
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuInsertDummyText 
         Caption         =   "&Dummy Text"
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&The Quick Brown Fox.."
            Index           =   0
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&Jackdaws..Quartz"
            Index           =   1
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&How Razorback-jumping Frogs.."
            Index           =   2
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&Cozy Lummox.."
            Index           =   3
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "Lorem Ipsum.."
            Index           =   4
         End
      End
      Begin VB.Menu mnuInsertKS 
         Caption         =   "&Keyboard Symbols"
      End
      Begin VB.Menu mnuInsertCharacter 
         Caption         =   "&Special Character..."
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuInsertUSymbol 
         Caption         =   "&Unicode..."
      End
      Begin VB.Menu mnuInsertLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertCitation 
         Caption         =   "&Citation..."
      End
      Begin VB.Menu mnuInsertHTMLXML 
         Caption         =   "&HTML/XML..."
         Begin VB.Menu mnuInsertStartingHTML 
            Caption         =   "&Starting HTML"
         End
         Begin VB.Menu mnuInsertSGMLCurrentFontInfo 
            Caption         =   "&Current XHTML Font Information"
         End
         Begin VB.Menu mnuInsertJavaScript 
            Caption         =   "&JavaScript in HTML..."
         End
         Begin VB.Menu mnuInsertStartingXML 
            Caption         =   "&Starting XML"
         End
      End
      Begin VB.Menu mnuInsertOPBpath 
         Caption         =   "&OpenBook file path..."
      End
      Begin VB.Menu mnuInsertLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertAccent 
         Caption         =   "&Accent"
         Begin VB.Menu mnuInsertAccentAcute 
            Caption         =   "Acute (á)"
         End
         Begin VB.Menu mnuInsertAccentGrave 
            Caption         =   "Grave (à)"
         End
         Begin VB.Menu mnuInsertAccentTilde 
            Caption         =   "Tilde (ã)"
         End
         Begin VB.Menu mnuInsertAccentUmlaut 
            Caption         =   "Umlaut (ä)"
         End
         Begin VB.Menu mnuInsertAccentCedilla 
            Caption         =   "Cedilla (ç)"
         End
         Begin VB.Menu mnuInsertAccentCaret 
            Caption         =   "Circumflex (â)"
         End
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatToggleFont 
         Caption         =   "Switch Font"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFormatFontsInDocument 
         Caption         =   "Fonts in Document"
         Begin VB.Menu mnuFormatFontsInDocumentFont 
            Caption         =   "Font"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFormatReplaceFonts 
         Caption         =   "Replace Fonts..."
      End
      Begin VB.Menu mnuFormatLine1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFormatParagraph 
         Caption         =   "&Paragraph..."
      End
      Begin VB.Menu mnuFormatFontScript 
         Caption         =   "&Font script..."
         Begin VB.Menu mnuFormatSuperscript 
            Caption         =   "&Superscript"
         End
         Begin VB.Menu mnuFormatSubscript 
            Caption         =   "&Subscript"
         End
      End
      Begin VB.Menu mnuFormatFontCase 
         Caption         =   "&Font Case"
         Begin VB.Menu mnuFormatCaseUppercase 
            Caption         =   "&Uppercase"
         End
         Begin VB.Menu mnuFormatCaseLowercase 
            Caption         =   "&Lowercase"
         End
         Begin VB.Menu mnuFormatCaseToggleCaps 
            Caption         =   "&Toggle All Caps"
         End
      End
      Begin VB.Menu mnuFormatIndentlm 
         Caption         =   "&Indent/Left Margin..."
         Begin VB.Menu mnuFormatIIndent 
            Caption         =   "&Increase Indent/Left Margin"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuFormatDIndent 
            Caption         =   "&Decrease Indent/Left Margin"
            Shortcut        =   ^{F2}
         End
      End
      Begin VB.Menu mnuFormatFontSize 
         Caption         =   "Font Size..."
         Begin VB.Menu mnuFormatChgSize 
            Caption         =   "&Increase Font Size"
            Index           =   0
            Shortcut        =   +{F9}
         End
         Begin VB.Menu mnuFormatChgSize 
            Caption         =   "&Decrease Font Size"
            Index           =   1
            Shortcut        =   +{F11}
         End
      End
      Begin VB.Menu mnuFormatCharOffset 
         Caption         =   "&Character Offset..."
         Begin VB.Menu mnuFormatSetCharOffset 
            Caption         =   "&Increase Character Offset"
            Index           =   0
            Shortcut        =   +{F5}
         End
         Begin VB.Menu mnuFormatSetCharOffset 
            Caption         =   "&Decrease Character Offset"
            Index           =   1
            Shortcut        =   +{F6}
         End
         Begin VB.Menu mnuFormatSetCharOffset 
            Caption         =   "&Set Character Offset"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFormatHangingIndent 
         Caption         =   "&Hanging Indent..."
         Begin VB.Menu mnuFormatIncreaseHangingIndent 
            Caption         =   "&Increase Hanging Indent"
         End
         Begin VB.Menu mnuFormatDecreaseHangingIndent 
            Caption         =   "&Decrease Hanging Indent"
         End
         Begin VB.Menu mnuFormatResetHangingIndent 
            Caption         =   "&Reset Hanging Indent"
         End
      End
      Begin VB.Menu mnuFormatLine2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFormatBullet 
         Caption         =   "&Switch Bullet Style"
      End
      Begin VB.Menu mnuFormatStyle 
         Caption         =   "&Style"
         Begin VB.Menu mnuFormatBold 
            Caption         =   "&Bold"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuFormatItalic 
            Caption         =   "&Italic"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuFormatUnderline 
            Caption         =   "&Underline"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuFormatstrikethru 
            Caption         =   "&Strikethru"
         End
      End
      Begin VB.Menu mnuFormatUnderlineStyle 
         Caption         =   "&Underline Style"
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Normal"
            Index           =   0
         End
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Dotted"
            Index           =   1
         End
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Dash"
            Index           =   2
         End
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Dot Dash"
            Index           =   3
         End
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Dot Dot Dash"
            Index           =   4
         End
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Thick"
            Index           =   5
         End
         Begin VB.Menu mnuFormatUnderlineStyleSub 
            Caption         =   "&Wave"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFormatAlignment 
         Caption         =   "&Alignment..."
         Begin VB.Menu mnuFormatAlignLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuFormatAlignCenter 
            Caption         =   "&Center"
         End
         Begin VB.Menu mnuFormatAlignRight 
            Caption         =   "&Right"
         End
         Begin VB.Menu mnuFormatAlignJustify 
            Caption         =   "&Justify"
         End
      End
      Begin VB.Menu mnuFormatHighlight 
         Caption         =   "&Highlight"
         Index           =   0
         Begin VB.Menu mnuFormatHighlight1 
            Caption         =   "&Highlight"
         End
         Begin VB.Menu mnuFormatHLUnHL 
            Caption         =   "Unhighlight"
         End
         Begin VB.Menu mnuFormatHC 
            Caption         =   "Highlight &Color..."
         End
      End
      Begin VB.Menu mnuFormatRealQuotes 
         Caption         =   "&Disable SymbolMatic"
      End
      Begin VB.Menu mnuFormatReplaceDQ 
         Caption         =   "Replace Straight Quotes"
      End
      Begin VB.Menu mnuFormatTabs 
         Caption         =   "&Tabs..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsDocStatistics 
         Caption         =   "&Document Statistics..."
      End
      Begin VB.Menu mnuToolsLiveWC 
         Caption         =   "&Live Word Count"
      End
      Begin VB.Menu mnuToolsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExtras 
         Caption         =   "&Extras"
         Begin VB.Menu mnuToolsExtrasReverse 
            Caption         =   "&Reverse Text"
         End
         Begin VB.Menu mnuToolsExtrasCJPunc 
            Caption         =   "Chinese/Japanese Mode"
         End
         Begin VB.Menu mnuToolsExtrasShowFrequency 
            Caption         =   "Count Occurrences..."
         End
      End
      Begin VB.Menu mnuToolsUnlimitMaxLength 
         Caption         =   "&Unlimit Maximum Length"
      End
      Begin VB.Menu mnuToolsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsMakeOpenBookFrom 
         Caption         =   "&Make OpenBook From"
         Begin VB.Menu mnuToolsMakeOpenBookFromRecentFiles 
            Caption         =   "&Recent Files"
         End
         Begin VB.Menu mnuToolsMakeOpenBookFromCurrentFiles 
            Caption         =   "&Currently Open Files"
         End
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "T&able"
      Begin VB.Menu mnuTableInsert 
         Caption         =   "&Insert..."
      End
      Begin VB.Menu mnuTableElastic 
         Caption         =   "&Elastic Table..."
      End
      Begin VB.Menu mnuTableLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTableAddColumn 
         Caption         =   "A&dd Column"
      End
      Begin VB.Menu mnuTableRemoveLastColumn 
         Caption         =   "&Remove Last Column"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowRestoreDown 
         Caption         =   "&Restore Current Window Down"
      End
      Begin VB.Menu mnuWindowRestoreUp 
         Caption         =   "&Restore Current Window Up"
      End
      Begin VB.Menu mnuWindowMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "&Minimize All"
      End
      Begin VB.Menu mnuWindowNext 
         Caption         =   "Next"
      End
      Begin VB.Menu mnuWindowLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
         Shortcut        =   +^{F7}
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
         Shortcut        =   +^{F8}
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About NIXON Hyperwrite"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "RightClick"
      Begin VB.Menu mnuRightClickCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuRightClickCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuRightClickPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuRightClickSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickFontsUsed 
         Caption         =   "Fonts Used in Document"
         Begin VB.Menu mnuRightClickFontsUsedFont 
            Caption         =   "Font"
            Index           =   0
         End
      End
      Begin VB.Menu mnuRightClickParagraph 
         Caption         =   "Paragraph..."
      End
      Begin VB.Menu mnuRightClickSwitchBullet 
         Caption         =   "Switch Bullet Style"
      End
      Begin VB.Menu mnuRightClickSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickGetInfo 
         Caption         =   "Get Info..."
      End
   End
End
Attribute VB_Name = "frmMain"
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
'External Functions
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
            (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
            ByVal lpsz2 As String) As Long
            
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
'Const TBSTYLE_FLAT As Long = &H800
'Const TB_SETSTYLE = WM_USER + 56
'Const TB_GETSTYLE = WM_USER + 57
'Const EM_LINEINDEX = &HBB
'Const EM_UNDO = &HC7
'Const EM_REDO = (&H400 + 84)
Const sQuote As String = """"
Dim lngHighlightColor As Long

'Color chooser without OCX
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias _
     "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Dim CustomColors() As Byte
'Font
Dim vFontFace(4) As String
Dim cFontColor(4) As Long
Dim intFontSize(4) As Integer
Dim FontBold(4) As Boolean
Dim FontItalic(4) As Boolean
Dim FontUnderline(4) As Boolean
Dim FontStrikethru(4) As Boolean
Dim btFontIndex(4) As Byte
Dim btFont As Byte

'Find/Replace
Dim lngCurrentPoint As Long

'Undo Types
'Private Const EM_GETUNDONAME = (WM_USER + 86)
Private Enum ERECUndoTypeConstants
    ercUID_UNKNOWN = 0
    ercUID_TYPING = 1
    ercUID_DELETE = 2
    ercUID_DRAGDROP = 3
    ercUID_CUT = 4
    ercUID_PASTE = 5
End Enum

Private Sub cboFontFace_KeyPress(KeyAscii As Integer)
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
If KeyAscii = 13 Then
    ActiveForm.rtfText.SelFontName = cboFontFace.Text
    cboFontFace_Click
End If
End Sub

Private Sub cmdFindPrev_Click()
On Error Resume Next
Dim lngPos As Long
'MsgBox InStr(ActiveForm.rtfText.Text, txtFind.Text) - Len(txtFind.Text)
'MsgBox ActiveForm.rtfText.SelStart
If InStr(ActiveForm.rtfText.Text, txtFind.Text) > ActiveForm.rtfText.SelStart Then
    lngPos = InStrRev(ActiveForm.rtfText.Text, txtFind.Text, Len(ActiveForm.rtfText.Text))
Else
    lngPos = InStrRev(ActiveForm.rtfText.Text, txtFind.Text, ActiveForm.rtfText.SelStart)
End If
If lngPos = 0 Then
    FlashStatus "Could not find " & Chr(147) & TrimLongWords(txtFind.Text, 10) & Chr(148)
    Exit Sub
End If
ActiveForm.rtfText.SelStart = lngPos - 1
ActiveForm.rtfText.SelLength = Len(txtFind.Text)
End Sub

Private Sub lblColor_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblColor.BackColor = &HFFFFFF
DoEvents
End Sub

Private Sub lblColor_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblColor.BackColor = cFontColor(btFont)
End Sub

Private Sub MDIForm_Activate()
DoToolbars (True)
End Sub

Private Sub MDIForm_DblClick()
LoadNewDoc
End Sub
 Private Function TranslateUndoType(ByVal eType As ERECUndoTypeConstants) As String
   Select Case eType
   Case 0 'Unknown
      TranslateUndoType = ""
   Case 1 'Typing
      TranslateUndoType = "Typing"
   Case 2 'Delete
      TranslateUndoType = "Delete"
   Case 3 'Drag/Drop
      TranslateUndoType = "Drag/Drop"
   Case 4 'Cut
      TranslateUndoType = "Cut"
   Case 5 'Paste
      TranslateUndoType = "Paste"
   End Select
End Function
Private Property Get UndoType() As ERECUndoTypeConstants
    Const EM_GETUNDONAME = (WM_USER + 86)
    UndoType = SendMessageLong(ActiveForm.rtfText.hwnd, EM_GETUNDONAME, 0, 0)
End Property

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bLiveWC = False Then StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub mnuEditUndoReplace_Click()
If ActiveForm.rtfText.Tag = "" Then Exit Sub
ActiveForm.rtfText.TextRTF = ActiveForm.rtfText.Tag
ActiveForm.rtfText.Tag = ""
mnuEditUndoReplace.Enabled = False
End Sub

Private Sub mnuFile_Click()
On Error Resume Next
DoMenus
If ActiveForm Is Nothing Then Exit Sub
mnuFileSave.Enabled = ActiveForm.bChanged = True
mnuFileRevert.Enabled = ActiveForm.bChanged = True
mnuFileSaveSelection.Enabled = ActiveForm.rtfText.SelLength <> 0
mnuFileRevert.Enabled = ActiveForm.rtfText.FileName <> ""
mnuFileAutoSave.Checked = ActiveForm.bAutoSave
End Sub
Private Sub DoMenus()
Dim bForm As Boolean
bForm = Not (ActiveForm Is Nothing)
mnuFileClose.Enabled = bForm
mnuFileCloseAll.Enabled = bForm
mnuFilePrint.Enabled = bForm
mnuFileRevert.Enabled = bForm
mnuFileAutoSave.Enabled = bForm
mnuFileSave.Enabled = bForm
mnuFileSaveAll.Enabled = bForm
mnuFileSaveAs.Enabled = bForm
mnuFileSaveSelection.Enabled = bForm
mnuEditAppend.Enabled = bForm
mnuEditChgProtection.Enabled = bForm
mnuEditClear.Enabled = bForm
mnuEditCopy.Enabled = bForm
mnuEditCopyList.Enabled = bForm
mnuEditCut.Enabled = bForm
mnuEditDelNextWord.Enabled = bForm
mnuEditDelPrevWord.Enabled = bForm
mnuEditFindNext.Enabled = bForm
'mnuEditFindReplace.Enabled = bForm
mnuEditGoTo.Enabled = bForm
mnuEditGoToBookmark.Enabled = bForm
mnuEditIncrementalFind.Enabled = bForm
mnuEditInsBookmark.Enabled = bForm
mnuEditLineSelect.Enabled = bForm
mnuEditPaste.Enabled = bForm
mnuEditPasteList.Enabled = bForm
mnuEditPastePlain.Enabled = bForm
mnuEditRedo.Enabled = bForm
mnuEditUndoReplace.Enabled = bForm
mnuEditSelAftCur.Enabled = bForm
mnuEditSelBefCur.Enabled = bForm
mnuEditSelectAll.Enabled = bForm
mnuEditSelectNextWord.Enabled = bForm
mnuEditSelectNone.Enabled = bForm
mnuEditSelectPrevWord.Enabled = bForm
mnuEditSelUpTo.Enabled = bForm
mnuEditTrimSpaces.Enabled = bForm
mnuEditUndo.Enabled = bForm
mnuViewPageWidth.Enabled = bForm
mnuViewRTF.Enabled = bForm
mnuViewTogglePrntNrml.Enabled = bForm
mnuGoTo.Enabled = bForm
mnuGoToEnd.Enabled = bForm
mnuGoToLine.Enabled = bForm
mnuGoToLineAbove.Enabled = bForm
mnuGoToLineBelow.Enabled = bForm
mnuGoToNextWord.Enabled = bForm
mnuGoToPrevWord.Enabled = bForm
mnuInsertAccent.Enabled = bForm
mnuInsertCharacter.Enabled = bForm
mnuInsertCitation.Enabled = bForm
mnuInsertDate.Enabled = bForm
mnuInsertDateandTime.Enabled = bForm
mnuInsertDateTime.Enabled = bForm
mnuInsertDummyText.Enabled = bForm
mnuInsertHTMLXML.Enabled = bForm
mnuInsertNonbreakingSpace.Enabled = bForm
mnuInsertObject.Enabled = bForm
mnuInsertOPBpath.Enabled = bForm
mnuInsertSGMLCurrentFontInfo.Enabled = bForm
mnuInsertTime.Enabled = bForm
mnuInsertUSymbol.Enabled = bForm
mnuInsertKS.Enabled = bForm
mnuFormatUnderlineStyle.Enabled = bForm
mnuFormatReplaceFonts.Enabled = bForm
mnuFormatFontsInDocument.Enabled = bForm
mnuFormatAlignment.Enabled = bForm
mnuFormatBullet.Enabled = bForm
mnuFormatCaseLowercase.Enabled = bForm
mnuFormatCaseUppercase.Enabled = bForm
mnuFormatCaseToggleCaps.Enabled = bForm
mnuFormatCharOffset.Enabled = bForm
mnuFormatDecreaseHangingIndent.Enabled = bForm
mnuFormatDIndent.Enabled = bForm
mnuFormatFontCase.Enabled = bForm
mnuFormatFontScript.Enabled = bForm
mnuFormatFontSize.Enabled = bForm
mnuFormatHangingIndent.Enabled = bForm
mnuFormatHC.Enabled = bForm
mnuFormatHighlight(0).Enabled = bForm
'mnuFormatIIndent.Enabled = bForm
mnuFormatIncreaseHangingIndent.Enabled = bForm
mnuFormatIndentlm.Enabled = bForm
'mnuFormatItalic.Enabled = bForm
mnuFormatParagraph.Enabled = bForm
mnuFormatRealQuotes.Enabled = bForm
mnuFormatReplaceDQ.Enabled = bForm
mnuFormatResetHangingIndent.Enabled = bForm
mnuFormatSetCharOffset(0).Enabled = bForm
mnuFormatstrikethru.Enabled = bForm
mnuFormatStyle.Enabled = bForm
mnuFormatTabs.Enabled = bForm
mnuFormatToggleFont.Enabled = bForm
'mnuFormatToggleSmartQuotes.Enabled = bForm
mnuToolsDocStatistics.Enabled = bForm
mnuToolsExtras.Enabled = bForm
mnuToolsLiveWC.Enabled = bForm
mnuToolsMakeOpenBookFromCurrentFiles.Enabled = bForm
mnuToolsUnlimitMaxLength.Enabled = bForm
mnuTableAddColumn.Enabled = bForm
mnuTableElastic.Enabled = bForm
mnuTableInsert.Enabled = bForm
mnuTableRemoveLastColumn.Enabled = bForm
End Sub

Private Sub mnuFileNewFromClipboard_Click()
LoadNewDoc
StatusBar.Panels(1).Text = "Importing Clipboard..."
mnuEditPaste_Click
StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub mnuFileSetASDuration_Click()

End Sub

Private Sub mnuFormat_Click()
On Error Resume Next
DoMenus
If InStr(ActiveForm.rtfText.SelRTF, "{\pict") <> 0 Or _
    InStr(ActiveForm.rtfText.SelRTF, "{\object") <> 0 Or _
    DoPrefs(0, "ParseFontTable") = 0 Then
    mnuFormatFontsInDocument.Enabled = False
    Exit Sub
Else
    mnuFormatFontsInDocument.Enabled = True
End If
StatusBar.Style = sbrSimple
StatusBar.SimpleText = "Getting Fonts in Document..."
Dim x As Integer
'Dim intFontPos As Integer, intFontEndPos As Integer
If mnuFormatFontsInDocumentFont.Count <> 1 Then
    For x = 1 To mnuFormatFontsInDocumentFont.Count
        If x <> 1 Then Unload mnuFormatFontsInDocumentFont(x - 1)
    Next
End If
For x = 0 To GetLastFontNum
    If x >= mnuFormatFontsInDocumentFont.Count Then Load mnuFormatFontsInDocumentFont(x)
    mnuFormatFontsInDocumentFont(x).Caption = ParseFontTable(x)
Next
StatusBar.Style = sbrNormal
StatusBar.SimpleText = "Ready"
If bRealSymbols = True Then
    mnuFormatRealQuotes.Caption = "&Disable SymbolMatic"
Else
    mnuFormatRealQuotes.Caption = "&Enable SymbolMatic"
End If
mnuFormatUnderlineStyle.Enabled = InStr(1, ActiveForm.rtfText.SelRTF, "\ul") <> 0
mnuFormatBold.Checked = ActiveForm.rtfText.SelBold
mnuFormatItalic.Checked = ActiveForm.rtfText.SelItalic
mnuFormatUnderline.Checked = ActiveForm.rtfText.SelUnderline
mnuFormatstrikethru.Checked = ActiveForm.rtfText.SelStrikeThru
mnuFormatCaseUppercase.Enabled = ActiveForm.rtfText.SelText <> ""
mnuFormatCaseLowercase.Enabled = ActiveForm.rtfText.SelText <> ""
End Sub

Private Sub mnuFormatFontsInDocumentFont_Click(Index As Integer)
ActiveForm.rtfText.SelFontName = mnuFormatFontsInDocumentFont(Index).Caption
End Sub

Private Sub mnuFormatHighlight1_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
Dim sSelected As String
Dim lCurrentStart As Long
sSelected = ActiveForm.rtfText.SelText
lCurrentStart = ActiveForm.rtfText.SelStart
Dim RTFformat As CHARFORMAT2
        RTFformat.cbSize = Len(RTFformat)
        RTFformat.dwMask = CFM_BACKCOLOR
        If lngHighlightColor = 0 Then
            RTFformat.crBackColor = vbYellow
        Else
            RTFformat.crBackColor = lngHighlightColor
        End If
    SendMessage ActiveForm.rtfText.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
End Sub

Private Sub mnuFormatReplaceFonts_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
frmReplaceFonts.Show vbModal, Me
End Sub

Private Sub mnuFormatSetCharOffset_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        ActiveForm.rtfText.SelCharOffset = ActiveForm.rtfText.SelCharOffset + 15
    Case 1
        ActiveForm.rtfText.SelCharOffset = ActiveForm.rtfText.SelCharOffset - 15
    Case 2
        ActiveForm.rtfText.SelCharOffset = InputBox("Enter Offset:", "Set Character Offset", ActiveForm.rtfText.SelCharOffset)
End Select
End Sub

Private Sub mnuFormatUnderlineStyleSub_Click(Index As Integer)
Dim strRTF As String, strUl As String
Dim lngStart As Long, lngLen As Long
lngStart = ActiveForm.rtfText.SelStart
lngLen = ActiveForm.rtfText.SelLength
Select Case Index
    Case 0 'Normal
        strUl = ""
    Case 1 'Dot
        strUl = "d"
    Case 2 'Dash
        strUl = "dash"
    Case 3 'Dot dash
        strUl = "dashd"
    Case 4 'Dot dot dash
        strUl = "dashdd"
    Case 5 'Thick
        strUl = "th"
    Case 6 'Wave
        strUl = "wave"
End Select
strRTF = Replace(ActiveForm.rtfText.SelRTF, "\uldashdd", "\ul")
strRTF = Replace(strRTF, "\uldashd", "\ul")
strRTF = Replace(strRTF, "\uldash", "\ul")
strRTF = Replace(strRTF, "\uldb", "\ul")
strRTF = Replace(strRTF, "\uld", "\ul")
strRTF = Replace(strRTF, "\ulth", "\ul")
strRTF = Replace(strRTF, "\ulwave", "\ul")
strRTF = Replace(strRTF, "\ulword", "\ul")
strRTF = Replace(strRTF, "\ul\", "\ul" & strUl & "\")
strRTF = Replace(strRTF, "\ul ", "\ul" & strUl & " ")
ActiveForm.rtfText.SelRTF = strRTF
ActiveForm.rtfText.SelStart = lngStart
ActiveForm.rtfText.SelLength = lngLen
End Sub

Private Sub mnuGo_Click()
DoMenus
End Sub

Private Sub mnuInsert_Click()
DoMenus
End Sub

Private Sub mnuInsertNonbreakingSpace_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SelText = Chr(160)
End Sub

Private Sub mnuInsertObject_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
InsertObj (ShowCommonDlg(True, "", Me, _
"Images (*.gif, *.jpg, *.bmp, *.dib, *.wmf, *.emf, *.ico, *.cur)" & Chr(0) & _
"*.gif;*.jpg;*.bmp;*.dib;*.wmf;*.emf;*.ico;*.cur" & Chr(0) & "All Files (*)" & Chr(0) & "*", "Insert Object", 4096))
10:
ErrorTrap "inserting a picture"
End Sub
Private Function InsertObj(sFile As String)
On Error GoTo 10
    StatusBar.Panels(1).Text = "Clearing Clipboard..."
    Clipboard.Clear
    StatusBar.Panels(1).Text = "Setting Data on Clipboard..."
    Clipboard.SetData LoadPicture(sFile)
    StatusBar.Panels(1).Text = "Getting Data from Clipboard..."
    SendMessage ActiveForm.rtfText.hwnd, WM_PASTE, 0&, ByVal 0&
    StatusBar.Panels(1).Text = "Ready"
10:
    If Err.Number = 481 Then
    CustomBox "Could not insert picture because it was invalid.", "You tried to insert an unsupported or invalid type of picture into your document.", vbExclamation, "", "", "OK"
    Exit Function
    End If
    ErrorTrap "inserting an object"
End Function

Private Function FindFunc(FStr As String, RTFBox As RichTextBox, bButton As Boolean)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Function
    End If
On Error GoTo 10
Static StartPoint As Long
Static inFind As Boolean
Dim lngPos As Long
Dim optionint As Integer
    If chkOptions.Item(0).Value = Checked Then
        optionint = optionint + rtfWholeWord
    End If
    If chkOptions.Item(1).Value = Checked Then
        optionint = optionint + rtfMatchCase
    End If
    If FStr = "" Then Exit Function
    If inFind = True Then lngCurrentPoint = RTFBox.SelStart
    If chkOptions(2).Value = Checked Then
        If bButton = True Then
            lngPos = myFind(StartPoint, lngCurrentPoint)
        Else
            lngPos = myFind(RTFBox.SelStart, 0)
        End If
    Else
        If bButton = True Then
            lngPos = myFind(RTFBox.SelStart + 1, RTFBox.SelStart)
        Else
            lngPos = myFind(RTFBox.SelStart, RTFBox.SelStart)
        End If
    End If
        If lngPos = -1 Then
            If Len(ActiveForm.rtfText.Text) = 0 Or InStr(1, ActiveForm.rtfText.Text, FStr, vbTextCompare) = 0 _
            Or chkOptions(3).Value = Unchecked Or chkOptions(1).Value = Checked Then
                FlashStatus "Could not find " & Chr(147) & TrimLongWords(FStr, 40) & Chr(148)
            Else
                lngPos = RTFBox.Find(FStr, 0, , optionint)
                StartPoint = 0
                inFind = True
                Exit Function
            End If
            StartPoint = 0
            lngCurrentPoint = 0
            inFind = True
        Else
            StartPoint = RTFBox.SelStart + RTFBox.SelLength
            inFind = False
        End If
10:
ErrorTrap "parsing a string"
End Function
Private Function FlashStatus(strText As String, Optional intLen As Integer = 6)
If StatusBar.Visible = True And DoPrefs(0, "StatusBarFind") = 1 Then
    Dim intSecond As Integer, bOriginalStatus As Boolean
    bOriginalStatus = StatusBar.Visible
    intSecond = Second(Now)
    StatusBar.Visible = True
    StatusBar.Style = sbrSimple
    StatusBar.SimpleText = strText
    tmrTimer.Enabled = True
Else
    CustomBox strText, "", vbInformation, "", "", "&OK"
End If
End Function
Private Function myFind(ByVal startP As Long, ByVal currP As Long) As Long
Dim lngPos As Long
Dim optionint As Integer
    If chkOptions.Item(0).Value = Checked Then
        optionint = optionint + rtfWholeWord
    End If
    If chkOptions.Item(1).Value = Checked Then
        optionint = optionint + rtfMatchCase
    End If
    lngPos = ActiveForm.rtfText.Find(txtFind.Text, startP, , optionint)
    myFind = lngPos
End Function
Private Function ReplaceFunc(txtRep As TextBox) As Boolean '(byRef currPoint As Integer) As Boolean
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Function
End If
Dim lngPos As Long
Dim StartPoint As Long
Static inFind As Boolean
StartPoint = ActiveForm.rtfText.SelStart
On Error GoTo 10
ActiveForm.rtfText.SetFocus
If txtFind.Text = "" Then
    ReplaceFunc = False
    Exit Function
End If
If inFind = True Then lngCurrentPoint = ActiveForm.rtfText.SelStart
    lngPos = myFind(StartPoint, lngCurrentPoint)
If lngPos = -1 Then
    StartPoint = 0
    lngCurrentPoint = 0
    inFind = True
    lngPos = myFind(StartPoint, lngCurrentPoint)
    If lngPos = -1 Then
        CustomBox "Hyperwrite has finished searching your document.", "", vbInformation, "", "", "OK"
        ReplaceFunc = False
    End If
Else
ActiveForm.rtfText.SelText = txtRep.Text
FindFunc txtFind.Text, ActiveForm.rtfText, False
ReplaceFunc = True
inFind = False
End If
10:
ErrorTrap "replacing text"
End Function

Private Sub cboFont_Click()
ChangeFont (cboFont.ListIndex)
ShowAttributes
End Sub

Private Sub cmdFindNext_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
FindFunc txtFind.Text, ActiveForm.rtfText, True
End Sub

Private Sub cmdReplace_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ReplaceFunc txtReplace '(0)
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
End Sub

Private Function ReplaceAll(strFind As String, strReplace As String, strReplaceAlt As String, bReport As Boolean)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Function
End If
On Error GoTo 10
Dim CountPoint As Long
Dim StartPoint As Long
Dim bReplaceNml As Boolean
Dim lngPos As Long
Dim optionint As Integer
Dim bFirstTime As Boolean
mnuEditUndoReplace.Tag = ActiveForm.rtfText.TextRTF
    If chkOptions.Item(0).Value = Checked Then
        optionint = optionint + 2
    End If
    If chkOptions.Item(1).Value = Checked Then
        optionint = optionint + 4
    End If
bNoStatus = True
ActiveForm.rtfText.SetFocus
If strFind = "" Then Exit Function
lngPos = 0
StartPoint = 0
CountPoint = 0
ActiveForm.rtfText.Tag = ActiveForm.rtfText.TextRTF
StatusBar.SimpleText = "Replacing..." & " Press ESC to cancel"
StatusBar.Style = sbrSimple
DoEvents
Do While lngPos <> -1
    If KeyDown(vbKeyEscape) Then Exit Do
        If bFirstTime = False Then
            lngPos = ActiveForm.rtfText.Find(strFind, StartPoint, , optionint)
            bFirstTime = True
        Else
            lngPos = ActiveForm.rtfText.Find(strFind, StartPoint + Len(strReplace), , optionint)
        End If
        StartPoint = ActiveForm.rtfText.SelStart
        If lngPos <> -1 Then
            If bReplaceNml = True Then
                ActiveForm.rtfText.SelText = strReplaceAlt
                bReplaceNml = False
            Else
                ActiveForm.rtfText.SelText = strReplace
                bReplaceNml = True
            End If
        CountPoint = CountPoint + 1
        End If
Loop
mnuEditUndoReplace.Enabled = True
StatusBar.Panels(1).Text = "Ready"
StatusBar.Style = sbrNormal
If lngPos = -1 Then
    If bReport = True Then
        If CountPoint > 1 Or CountPoint = 0 Then
        CustomBox "Hyperwrite has finished searching your document and " & CountPoint & " replacements were made.", "", vbInformation, "", "", "OK"
        End If
        If CountPoint = 1 Then
        CustomBox "Hyperwrite has finished searching your document and " & CountPoint & " replacement was made.", "", vbInformation, "", "", "OK"
        End If
    End If
    bNoStatus = False
    If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
End If
10:
ErrorTrap "replacing text (all)"
End Function



Private Sub cmdReplaceAll_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ReplaceAll txtFind.Text, txtReplace.Text, txtReplace.Text, True
End Sub

Private Sub cmdSimpleReplace_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
Dim intMsgReturn As Integer
intMsgReturn = CustomBox("Are you sure you want to use Quick Replace?", "Quick Replace does not preserve formatting in your document. You should only use it for long, plain text files.", vbExclamation, "", "&Cancel", "&Use")
If intMsgReturn = 1 Then ActiveForm.rtfText.Text = Replace(ActiveForm.rtfText.Text, txtFind.Text, txtReplace.Text)
10:
ErrorTrap "replacing text (all, quick)"
End Sub

Private Sub cboFontSize_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboFontSize_Click
End Sub

Private Sub mnuInsertSGMLCurrentFontInfo_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    On Error GoTo 10
    Dim strFace As String, strStyle As String, strDecoration As String
    strFace = ActiveForm.rtfText.SelFontSize & "pt " & "'" & ActiveForm.rtfText.SelFontName & "';"
    Select Case ActiveForm.rtfText.SelAlignment
        Case rtfLeft
            strStyle = "text-align:left;"
        Case rtfCenter
            strStyle = "text-align:center;"
        Case rtfRight
            strStyle = "text-align:right;"
    End Select
    If ActiveForm.rtfText.SelUnderline = True Then strDecoration = "underline"
    If ActiveForm.rtfText.SelStrikeThru = True Then strDecoration = "line-through"
    If ActiveForm.rtfText.SelUnderline = True And ActiveForm.rtfText.SelStrikeThru = True Then
        strDecoration = "underline line-through"
    End If
    If ActiveForm.rtfText.SelBold = True Then strStyle = strStyle & "font-weight:bold;"
    If ActiveForm.rtfText.SelItalic = True Then strStyle = strStyle & "font-style:italic;"
    If strDecoration <> "" Then strStyle = strStyle & "text-decoration:" & strDecoration & ";"
    ActiveForm.rtfText.SelText = "<span style=" & sQuote & "font: " & strFace & strStyle & sQuote & "></span>"
10:
End Sub


Private Sub mnuEditChgProtection_Click()
On Error Resume Next
Dim lngPos As Long, lngSel As Long
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
lngPos = ActiveForm.rtfText.SelStart
lngSel = ActiveForm.rtfText.SelLength
If ActiveForm.rtfText.SelLength = 0 Then mnuEditSelectAll_Click
Select Case ActiveForm.rtfText.SelProtected
Case True
    ActiveForm.rtfText.SelProtected = False
Case False
    ActiveForm.rtfText.SelProtected = True
Case Else
    ActiveForm.rtfText.SelProtected = False
End Select
ActiveForm.rtfText.SelStart = lngPos
ActiveForm.rtfText.SelLength = lngSel
End Sub

Private Sub mnuEditDelNextWord_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{DEL}"
End Sub

Private Sub mnuEditDelPrevWord_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{BKSP}"
End Sub

Private Sub mnuEditfindReplace_Click()
On Error Resume Next
    mnuEditFindReplace.Checked = Not mnuEditFindReplace.Checked
    pctFindReplace.Visible = mnuEditFindReplace.Checked
    If pctFindReplace.Visible = True Then
        txtFind.SetFocus
        If ActiveForm.rtfText.SelText > "" Then
            txtFind.Text = ActiveForm.rtfText.SelText
            txtFind.SelLength = Len(txtFind.Text)
            txtFind.SetFocus
        End If
        mnuEditFindNext.Enabled = True
        Toolbar.Buttons(6).Value = tbrPressed
    Else
        mnuEditFindNext.Enabled = False
        Toolbar.Buttons(6).Value = tbrUnpressed
    End If
End Sub



Private Sub mnuEditIncrementalFind_Click()
mnuEditfindReplace_Click
chkOptions(2).Value = Checked
If pctFindReplace.Visible = False Then
    chkOptions(2).Value = Unchecked
End If
End Sub

Private Sub mnuEditPreferences_Click()
frmPrefs.Show vbModal, Me
End Sub
Private Sub mnuEditSelectNextWord_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
DoWords False, True
End Sub
Private Sub mnuEditSelectPrevWord_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
DoWords True, True
End Sub

Private Sub mnuFormatAlignCenter_Click()
On Error Resume Next
ActiveForm.rtfText.SelAlignment = rtfCenter
End Sub

Private Sub mnuFormatAlignJustify_Click() 'Justify alignment is only available as roundtripping
On Error Resume Next
ActiveForm.rtfText.SetFocus
SendKeys "^J"
End Sub

Private Sub mnuFormatAlignLeft_Click()
On Error Resume Next
ActiveForm.rtfText.SelAlignment = rtfLeft
End Sub

Private Sub mnuFormatAlignRight_Click()
On Error Resume Next
ActiveForm.rtfText.SelAlignment = rtfRight
End Sub

Private Sub mnuFormatCaseToggleCaps_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^+A"
End Sub

Private Sub mnuFormatChgSize_Click(Index As Integer)
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
If Index = 0 Then
    ActiveForm.rtfText.SelFontSize = ActiveForm.rtfText.SelFontSize + 2
Else
    ActiveForm.rtfText.SelFontSize = ActiveForm.rtfText.SelFontSize - 2
End If
End Sub

Private Sub mnuFormatReplaceDQ_Click()
On Error Resume Next
Dim strRTF As String
ReplaceAll " '", Chr(145), Chr(145), False
ReplaceAll vbTab & "'", Chr(145), Chr(145), False
ReplaceAll vbNewLine & "'", vbNewLine & Chr(145), vbNewLine & Chr(145), False
ReplaceAll "'", Chr(146), Chr(146), False
ReplaceAll " " & sQuote, Chr(147), Chr(147), False
ReplaceAll vbTab & sQuote, Chr(147), Chr(147), False
ReplaceAll vbNewLine & sQuote, vbNewLine & Chr(147), vbNewLine & Chr(147), False
ReplaceAll sQuote, Chr(148), Chr(148), False
End Sub

Private Sub mnuFormatToggleFont_Click()
Static bFirstFont As Boolean
If btFont < 4 Then
    ChangeFont (CInt(btFont) + 1)
Else
    ChangeFont (0)
End If
End Sub

Private Sub mnuGoToLineAbove_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{UP}"
End Sub

Private Sub mnuGoToLineBelow_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{DOWN}"
End Sub

Private Sub mnuGoToNextWord_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{RIGHT}"
End Sub

Private Sub mnuGoToPrevWord_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{LEFT}"
End Sub

Private Sub mnuInsertAccentAcute_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
If bRealSymbols = True Then
    bRealSymbols = False 'Wait for SymbolMatic to disable; DoEvents will not work.
    SendKeys "^'", 1
    bRealSymbols = True
Else
    SendKeys "^'"
End If
FlashStatus "Type for an acute accent"
End Sub

Private Sub mnuInsertAccentCaret_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^+6"
FlashStatus "Type for a caret accent"
End Sub

Private Sub mnuInsertAccentCedilla_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^,"
FlashStatus "Type a C for a cedilla"
End Sub

Private Sub mnuInsertAccentGrave_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^`"
FlashStatus "Type for a grave accent"
End Sub

Private Sub mnuInsertAccentTilde_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^+`"
FlashStatus "Type for a tilde accent"
End Sub

Private Sub mnuInsertAccentUmlaut_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^;"
FlashStatus "Type for an umlaut accent"
End Sub
Private Sub HandleNoWindows()
FlashStatus "Could not complete your request because there are no windows open.", 12
End Sub
Private Sub mnuInsertSampleSentence_Click(Index As Integer)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Select Case Index
Case 0
ActiveForm.rtfText.SelText = "The quick brown fox jumps over the lazy dog. "
Case 1
ActiveForm.rtfText.SelText = "Jackdaws love my big sphinx of quartz. "
Case 2
ActiveForm.rtfText.SelText = "How razorback-jumping frogs can level six piqued gymnasts! "
Case 3
ActiveForm.rtfText.SelText = "Cozy lummox gives smart squid who asks for job pen. "
Case 4
ActiveForm.rtfText.SelText = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. "
End Select
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
10:
ErrorTrap "inserting dummy text"
End Sub



Private Sub mnuOpenBook_Click()
OpenBook (ShowCommonDlg(True, "", Me, "OpenBook file (*.opb)" & Chr(0) & "*.opb" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Open Book", 4096))
End Sub
Private Function OpenBook(sFile As String)
On Error GoTo 10
If sFile = "" Then Exit Function
Dim CurrStart As Long
Dim EndStart As Long
Dim StrLength As Long
Dim FileName$
Dim FileNum%
Dim lngPos1 As Long
Dim lngPos2 As Long
Dim strBook As String
Dim bFirstTime As Boolean
If ActiveForm Is Nothing Then LoadNewDoc
strBook = OpenBinary(sFile)
lngPos1 = -1
Do While lngPos1 <> 0
    If bFirstTime = False Then
        lngPos1 = InStr(lngPos1 + 2, strBook, "<")
        bFirstTime = True
    Else
        lngPos1 = InStr(lngPos1 + 1, strBook, "<")
    End If
    lngPos2 = InStr(lngPos2 + 1, strBook, ">")
    FileName$ = Mid$(strBook, lngPos1 + 1, lngPos2 - lngPos1 - 1)
    btDocumentCount = btDocumentCount + 1
    LoadNewDoc
    OpenFile FileName$, False, , True
Loop
10:
End Function

Private Sub mnuGoTo_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim GStr As String
GStr = InputBox("Go to character...", "Go To...")
ActiveForm.SetFocus
ActiveForm.rtfText.SelStart = GStr
10:
End Sub

Private Sub mnuGoToLine_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim lngStart As Long
Dim GStr As Long
GStr = InputBox("Go to line#...", "Go To Line#...")
GoLine (GStr)
10:
End Sub
Private Function GoLine(GLng As Long)
On Error GoTo 10
Const EM_LINEINDEX = &HBB
Dim lngStart As Long
ActiveForm.rtfText.SetFocus
lngStart = SendMessage(ActiveForm.rtfText.hwnd, EM_LINEINDEX, GLng - 1, 0&)
ActiveForm.rtfText.SelStart = lngStart 'Go To line
Exit Function
10:
CustomBox "Invalid line number", "Please do not enter any symbols before/after the line number.", vbExclamation, "", "", "OK"
ActiveForm.rtfText.SetFocus
End Function

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Dim nForm As Form
    For Each nForm In Forms
        If Not (nForm.Name = "frmMain") Then
            Unload nForm
        End If
    Next
    If DoPrefs(0, "SaveWorkspace") = 1 Then
        DoPrefs 1, "ShowCoolbar", IIf(cbCoolBar.Visible, 1, 0)
        DoPrefs 1, "ShowToolbar", IIf(Toolbar.Visible, 1, 0)
        DoPrefs 1, "ShowFormatBar", IIf(tbrFontFormat.Visible, 1, 0)
        DoPrefs 1, "ShowExtFormatting", IIf(tbFormat.Visible, 1, 0)
        DoPrefs 1, "ShowSymbolBar", IIf(tbSymbols.Visible, 1, 0)
        DoPrefs 1, "ShowStatusBar", IIf(StatusBar.Visible, 1, 0)
        DoPrefs 1, "ShowRuler", IIf(pcSlider.Visible, 1, 0)
        DoPrefs 1, "WindowState", Me.WindowState
    End If
    SavePrefFile
End Sub

Private Sub mnuEditAppend_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim AppendText As String
If Clipboard.GetText = "" Then Exit Sub
AppendText = Clipboard.GetText
Clipboard.Clear
AppendText = AppendText + ActiveForm.rtfText.SelText
Clipboard.SetText AppendText
End Sub

Private Sub mnuEditFindNext_Click()
cmdFindNext_Click
End Sub

Private Sub mnuEditGoToBookmark_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmBookmarks.Show vbModal, Me
10:
End Sub

Private Sub mnuEditInsBookmark_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
frmBookmarks.lstList.AddItem ActiveForm.rtfText.SelStart
10:
End Sub

Private Sub mnuEditLineSelect_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
SelectLine fMainForm.ActiveForm.rtfText, ActiveForm.rtfText.GetLineFromChar(ActiveForm.rtfText.SelStart) + 1
End Sub

Private Sub mnuEditPastePlain_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelText = Clipboard.GetText
StatusBar.Panels(1).Text = "Ready"
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
End Sub

Private Sub mnuEditPurge_Click()
Dim YesNo%
YesNo% = CustomBox("This action cannot be undone. Would you like to continue?", "This will erase all clipboard contents.", vbExclamation, "", "Cancel", "Purge")
If YesNo% = 1 Then Clipboard.Clear
End Sub





Private Sub mnuEditSelUpTo_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim UInt As Long
StatusBar.Panels(1).Text = "Selecting to..."
UInt = InputBox("Select up to...", "Select up to...")
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = UInt
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuEditSelBefCur_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim CurPosition As Long
StatusBar.Panels(1).Text = "Selecting to cursor..."
CurPosition = ActiveForm.rtfText.SelStart
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = CurPosition
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuEditSelAftCur_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim CurPosition As Long
StatusBar.Panels(1).Text = "Selecting to cursor..."
CurPosition = ActiveForm.rtfText.SelStart
ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText.Text) - CurPosition
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuGoToEnd_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SelStart = Len(ActiveForm.rtfText.Text)
10:
End Sub

Private Sub mnuFileOpenText_Click()
On Error GoTo 10
Dim sFile As String
Dim YesNoCancel%
    sFile = ShowCommonDlg(True, "", Me, "All Files (*)" & Chr(0) & "*" & Chr(0), "Open...", 4096)
    OpenFile sFile, True, rtfText
    StatusBar.Panels(1).Text = "Ready"
Exit Sub
10:
ErrorTrap "opening a file as text", ParseFileName(sFile)
End Sub

Private Sub mnuFileGetInfo_Click()
frmGetInfo.Show vbModal, Me
End Sub

Private Sub mnuFormatHC_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
lngHighlightColor = ShowColorDlg
If ActiveForm.rtfText.SelText <> "" Then mnuFormatHighlight1_Click
End Sub

Private Sub mnuFormatHLUnHL_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim sSelected As String
Dim lCurrentStart As Long
Dim lCurrentSel As Long
sSelected = ActiveForm.rtfText.SelText
lCurrentStart = ActiveForm.rtfText.SelStart
lCurrentSel = ActiveForm.rtfText.SelLength
ActiveForm.rtfText.SelText = "\h0" & ActiveForm.rtfText.SelText
ActiveForm.rtfText.TextRTF = Replace(ActiveForm.rtfText.TextRTF, "\\h0", "\highlight1" & ActiveForm.rtfText.SelText & "\highlight0 `")
ActiveForm.rtfText.SelStart = lCurrentStart
ActiveForm.rtfText.SelLength = 1
ActiveForm.rtfText.SelText = ""
ActiveForm.rtfText.SelStart = lCurrentStart
ActiveForm.rtfText.SelLength = lCurrentSel
End Sub
Private Sub mnuFormatParagraph_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmFormat.Show vbModal, Me
End Sub
Private Sub CharOffset(btOption As Byte)
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
If btOption = 0 Then ActiveForm.rtfText.SelCharOffset = 0
If btOption = 1 Then ActiveForm.rtfText.SelCharOffset = ActiveForm.rtfText.SelCharOffset + 3
If btOption = 2 Then ActiveForm.rtfText.SelCharOffset = ActiveForm.rtfText.SelCharOffset - 3
End Sub
Private Sub mnuFormatRealQuotes_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
bRealSymbols = Not (bRealSymbols)
If bRealSymbols = True Then mnuFormatRealQuotes.Caption = "Disable SymbolMatic"
If bRealSymbols = False Then mnuFormatRealQuotes.Caption = "Enable SymbolMatic"
End Sub

Private Sub mnuFormatTabs_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmTabStops.Show vbModal, Me
End Sub

Private Sub mnuInsertCitation_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
frmCitation.Show vbModal, Me
10:
ErrorTrap "attempting to show insert citation dialog"
End Sub


Private Sub mnuInsertKS_Click()
On Error GoTo 10
Dim i As Integer
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
bNoStatus = True
For i = 33 To 126
ActiveForm.rtfText.SelText = Chr(i)
Next
bNoStatus = False
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
10:
ErrorTrap "inserting keyboard symbols"
End Sub


Private Sub mnuInsertOPBpath_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim sFile As String
sFile = ShowCommonDlg(True, "", Me, "Text Documents (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr(0) & "*.rtf;*.wri;*.doc;*.txt;*.text" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Insert OpenBook Path...", 4096)
If sFile = "" Then Exit Sub
ActiveForm.rtfText.SelText = "<" & Trim(sFile)
ActiveForm.rtfText.SelText = ">"
10:
End Sub

Private Sub mnuInsertUSymbol_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error Resume Next
Dim UStr As String
UStr = InputBox("Enter the character code for the symbol you want to insert.", "Insert Unicode")
With ActiveForm.rtfText
    .SelText = "\u" & .SelText & "?"
    .TextRTF = Replace(.TextRTF, "\\u" & .SelText & "?", "\u" & UStr & "?")
End With
End Sub

Private Sub mnuRightClick_Click()
On Error Resume Next
Dim x As Integer, strFontTable As String
If InStr(ActiveForm.rtfText.SelRTF, "{\pict") <> 0 Or _
    InStr(ActiveForm.rtfText.SelRTF, "{\object") <> 0 Or _
    DoPrefs(0, "ParseFontTable") = 0 Then
    mnuRightClickFontsUsed.Enabled = False
    Exit Sub
Else
    mnuRightClickFontsUsed.Enabled = True
End If
StatusBar.Style = sbrSimple
StatusBar.SimpleText = "Getting Fonts in Document... Press Esc to cancel"
If mnuRightClickFontsUsedFont.Count <> 1 Then
    For x = 1 To mnuRightClickFontsUsedFont.Count
        If x <> 1 Then Unload mnuRightClickFontsUsedFont(x - 1)
    Next
End If
For x = 0 To GetLastFontNum
    If x >= mnuRightClickFontsUsedFont.Count Then Load mnuRightClickFontsUsedFont(x)
    mnuRightClickFontsUsedFont(x).Caption = ParseFontTable(x)
    mnuRightClickFontsUsedFont(x).Visible = True
Next
StatusBar.Style = sbrNormal
StatusBar.SimpleText = "Ready"
mnuRightClickCut.Enabled = ActiveForm.rtfText.SelLength > 0
mnuRightClickCopy.Enabled = ActiveForm.rtfText.SelLength > 0
End Sub

Private Sub mnuRightClickCopy_Click()
mnuEditCopy_Click
End Sub

Private Sub mnuRightClickCut_Click()
mnuEditCut_Click
End Sub

Private Sub mnuRightClickFontsUsedFont_Click(Index As Integer)
ActiveForm.rtfText.SelFontName = mnuRightClickFontsUsedFont(Index).Caption
End Sub

Private Sub mnuRightClickGetInfo_Click()
With ActiveForm
    If .rtfText.SelText <> "" Then
        If InStr(1, .rtfText.SelRTF, "{\pict") <> 0 Then
            Dim lngPicH As Long, lngPicW As Long
            Dim lngPicHEnd As Long, lngPicWEnd As Long
            Dim lngPicStart As Long, lngPicEnd As Long, strSize As String
            Dim strDimensions As String
            lngPicW = InStr(1, .rtfText.SelRTF, "\picwgoal")
            lngPicH = InStr(1, .rtfText.SelRTF, "\pichgoal")
            lngPicWEnd = InStr(lngPicW + 1, .rtfText.SelRTF, "\")
            lngPicHEnd = InStr(lngPicH + 1, .rtfText.SelRTF, vbCr)
            strDimensions = "Width: " & CLng(Mid(.rtfText.SelRTF, lngPicW + 9, lngPicWEnd - lngPicW - 9) / 15) _
                & "px  Height: " & CLng(Mid(.rtfText.SelRTF, lngPicH + 9, lngPicHEnd - lngPicH - 9) / 15) & "px"
            lngPicStart = InStr(1, .rtfText.SelRTF, "{\pict")
            lngPicEnd = InStr(lngPicStart, .rtfText.SelRTF, "}")
            strSize = LenB(Mid(.rtfText.SelRTF, InStr(lngPicStart, .rtfText.SelRTF, vbCr), lngPicEnd - lngPicStart)) - 16
            CustomBox "Picture Info", strDimensions & "  Size: " & strSize & " bytes", vbInformation, "", "", "&OK"
        Else
            CustomBox "Selection Info", "Length: " & .rtfText.SelLength & "  Starting Position: " & _
            .rtfText.SelStart & "  Line: " & .rtfText.GetLineFromChar(.rtfText.SelStart) + 1 & vbNewLine & _
            "Standalone bytes: " & LenB(.rtfText.SelText) & " (Includes " & LenB(.rtfText.SelRTF) & "/" & _
            LenB(.rtfText.TextRTF) & ")", vbInformation, "", "", "&OK"
        End If
    Else
        mnuFileGetInfo_Click
    End If
End With
End Sub

Private Sub mnuRightClickParagraph_Click()
mnuFormatParagraph_Click
End Sub

Private Sub mnuRightClickPaste_Click()
mnuEditPaste_Click
End Sub

Private Sub mnuRightClickSwitchBullet_Click()
mnuFormatBullet_Click
End Sub

Private Sub mnuTable_Click()
DoMenus
End Sub

Private Sub mnuTableAddColumn_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
Dim lngPos As Long
Dim lngSlashPos As Long
Dim strRTF As String
Dim lngCellWidthPos As Long
Dim lngCellSlashPos As Long
Dim lngCellWidth As Long
Dim strAfterStuff As String
Dim strLeft As String
Dim lngOccurences As Long
strRTF = ActiveForm.rtfText.SelRTF
lngPos = InStrRev(strRTF, "\cellx")
lngSlashPos = InStr(lngPos + 1, strRTF, "\")
strAfterStuff = Right$(strRTF, Len(strRTF) - lngSlashPos + 1)
strAfterStuff = Replace(strAfterStuff, "\cell\row", "\cell\cell\row")
lngCellWidthPos = InStr(1, strRTF, "\cellx")
lngCellSlashPos = InStr(lngCellWidthPos + 1, strRTF, "\")
lngOccurences = FindOccurrences(strRTF, "\cellx") + 2
lngCellWidth = CLng(Mid$(strRTF, lngCellWidthPos + 6, lngCellSlashPos - lngCellWidthPos - 6)) * lngOccurences
strLeft = Left$(strRTF, lngSlashPos - 1)
ActiveForm.rtfText.SelRTF = strLeft & "\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs\cellx" & lngCellWidth & strAfterStuff
10:
End Sub


Private Sub mnuTableElastic_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
bRubberBand = Not (bRubberBand)
mnuTableElastic.Checked = bRubberBand
If bRubberBand = False Then ActiveForm.txtdrag.Visible = False
End Sub

Private Sub mnuTableInsert_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmTables.Show vbModal, Me
End Sub

Private Sub mnuTableRemoveLastColumn_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
Dim lngStartPos As Long, lngCellPos As Long, lngSlashPos As Long, lngWidthPos As Long
Dim strRTF As String, strLeft As String, strRight As String
lngStartPos = InStrRev(ActiveForm.rtfText.SelRTF, "\clbrdrt")
lngCellPos = InStrRev(ActiveForm.rtfText.SelRTF, "\cellx")
lngSlashPos = InStr(lngCellPos + 1, ActiveForm.rtfText.SelRTF, "\")
lngWidthPos = Mid$(ActiveForm.rtfText.SelRTF, lngCellPos + 6, lngSlashPos - lngCellPos - 6)
strLeft = Left$(ActiveForm.rtfText.SelRTF, lngStartPos)
strRight = Replace(Right$(ActiveForm.rtfText.SelRTF, Len(ActiveForm.rtfText.SelRTF) - lngSlashPos), "\cell\cell\row", "\cell\row", , 1) '"\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs\cellx" & lngWidthPos, "", lngStartPos, 1)
ActiveForm.rtfText.SelRTF = strLeft + strRight
10:
End Sub

Private Sub mnuTools_Click()
DoMenus
On Error Resume Next
If ActiveForm.rtfText.MaxLength = 0 Then
    mnuToolsUnlimitMaxLength.Caption = "&Limit Maximum Length"
Else
    mnuToolsUnlimitMaxLength.Caption = "&Unlimit Maximum Length"
End If
End Sub

Private Sub mnuInsertStartingHTML_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim HTMLStr As String
HTMLStr = "<!--Place DOCTYPE declaration here-->" + vbNewLine + "<html>" + vbNewLine + "<head>" + vbNewLine + "<title>Title</title>" + vbNewLine + "<meta name=" + sQuote + "keywords" + sQuote + " content=" + sQuote + sQuote + " />" + vbNewLine + "<meta name=" + sQuote + "description" + sQuote + " content=" + sQuote + sQuote + " />" + vbNewLine + "</head>" + vbNewLine + "<body>" + vbNewLine + "<div>" + vbNewLine + "</div>" + vbNewLine + "</body>" + vbNewLine + "</html>"
ActiveForm.rtfText.SelText = HTMLStr
10:
End Sub

Private Sub mnuToolsExtrasCJPunc_Click()
mnuToolsExtrasCJPunc.Checked = Not (mnuToolsExtrasCJPunc.Checked)
bAltMode = mnuToolsExtrasCJPunc.Checked
If bAltMode = True Then
    If CustomBox("Chinese/Japanese Punctuation Mode", "When this mode is on, you can type Chinese/Japanese punctuation marks from the keyboard.", vbInformation, "", "&More Info", "&OK") = 2 Then
        CustomBox "Shortcuts to use the Chinese/Japanese Punctuation Mode", "Ctrl and ; creates a left corner bracket, Ctrl and ' creates a right corner bracket , typing a period creates an ideographic full stop.", vbInformation, "", "", "&OK"
    End If
End If
End Sub

Private Sub mnuToolsExtrasReverse_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
If ActiveForm.rtfText.SelText = "" Then
    ActiveForm.rtfText.Text = StrReverse(ActiveForm.rtfText.Text)
Else
    ActiveForm.rtfText.SelText = StrReverse(ActiveForm.rtfText.SelText)
End If
Exit Sub
10:
ErrorTrap "reversing text"
End Sub

Private Sub mnuToolsExtrasShowFrequency_Click()
On Error Resume Next
Dim strWord As String, intOccurs As Integer
strWord = InputBox("String:", "Show Frequency")
If strWord <> "" Then 'Prevent loop at FindOccurences
    intOccurs = FindOccurrences(vbNullChar & ActiveForm.rtfText.Text, strWord)
Else
    intOccurs = -1
End If
If intOccurs = -1 Then
    CustomBox "The string " & Chr(147) & strWord & Chr(148) & " occurs 0 times in the current document.", "", vbInformation, "", "", "&OK"
Else
    CustomBox "The string " & Chr(147) & strWord & Chr(148) & " appears " & _
     intOccurs + 1 & " times in the current document.", "", vbInformation, "", "", "&OK"
End If
End Sub

Private Sub mnuToolsLiveWC_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
bLiveWC = Not (bLiveWC)
If bLiveWC = True Then
    StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
Else
    StatusBar.Panels(1).Text = "Ready"
End If
End Sub

Private Sub mnuToolsMakeOpenBookFromCurrentFiles_Click()
Dim strBook As String
Dim sDlgFile As String
Dim i As Integer
    For i = 1 To Forms.Count - 1
        If ActiveForm.rtfText.FileName <> "" Then
            strBook = strBook & "<" & ActiveForm.rtfText.FileName & ">"
            SendKeys "^{F6}", 1
            DoEvents
        End If
    Next
If strBook = "" Then
    CustomBox "No OpenBook could be made.", "All of the files which are currently open have not been saved yet. Hyperwrite needs filenames to write to the OpenBook.", vbExclamation, "", "", "&OK"
    Exit Sub
End If
sDlgFile = ShowCommonDlg(False, "", Me, "OpenBook file (*.opb)" & Chr(0) & "*.opb" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Save OpenBook...", 0)
If sDlgFile = "" Then Exit Sub
Dim FileNum%
FileNum% = FreeFile
Open sDlgFile For Output As FileNum%
Print #FileNum%, strBook
Close #FileNum%
End Sub

Private Sub mnuToolsMakeOpenBookFromRecentFiles_Click()
Dim strFile As String
Dim sDlgFile As String
Dim i As Integer
For i = 0 To 4
    strFile = strFile & "<" & DoPrefs(0, "Recent" & i + 1) & ">"
Next
sDlgFile = ShowCommonDlg(False, "", Me, "OpenBook file (*.opb)" & Chr(0) & "*.opb" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Save OpenBook...", 0)
If sDlgFile = "" Then Exit Sub
Dim FileNum%
FileNum% = FreeFile
Open sDlgFile For Output As FileNum%
Print #FileNum%, strFile
Close #FileNum%
End Sub

Private Function TrimSymbols(strWord As String) As String
Dim i As Integer
For i = 65 To 90
    TrimSymbols = Replace(strWord, Chr(i), "")
Next
For i = 97 To 122
    TrimSymbols = Replace(strWord, Chr(i), "")
Next
TrimSymbols = Replace(strWord, Chr(160), "")
End Function

Private Sub mnuToolsUnlimitMaxLength_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
With ActiveForm.rtfText
If .MaxLength = 0 Then
    .MaxLength = 20000000
Else
    .MaxLength = 0
End If
End With
End Sub

Private Sub mnuView_Click()
mnuViewCoolbar.Checked = cbCoolBar.Visible
mnuViewToolbar.Checked = Toolbar.Visible
mnuViewFormatBar.Checked = tbrFontFormat.Visible
mnuViewExtendedFormatting.Checked = tbFormat.Visible
mnuViewAccentsBar.Checked = tbSymbols.Visible
mnuViewStatusBar.Checked = StatusBar.Visible
mnuViewRuler.Checked = pcSlider.Visible
DoEvents
DoMenus
End Sub

Private Sub mnuViewCoolbar_Click()
    cbCoolBar.Visible = Not (cbCoolBar.Visible)
    mnuViewCoolbar.Checked = cbCoolBar.Visible
End Sub

Private Sub mnuViewFormatBar_Click()
    tbrFontFormat.Visible = Not (tbrFontFormat.Visible)
    cbCoolBar.Bands(2).Visible = tbrFontFormat.Visible
    mnuViewFormatBar.Checked = tbrFontFormat.Visible
End Sub

Private Sub mnuViewPageWidth_Click()
frmPaperSizes.Show vbModal, Me
End Sub

Private Sub mnuViewRTFCode_Click(Index As Integer)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
If Index = 0 Then
    Me.mnuViewRTF.Tag = "Whole"
Else
    Me.mnuViewRTF.Tag = "Sel"
End If
frmRTFCode.Show vbModal, Me
End Sub

Private Sub mnuViewTogglePrntNrml_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
bNormal = Not (bNormal)
If bNormal = True Then
ActiveForm.rtfText.RightMargin = lRightMargin
Else
ActiveForm.rtfText.RightMargin = Printer.Width
End If
10:
ActiveForm.ScaleMode = vbInches
ActiveForm.rtfText.RightMargin = 8.5
ActiveForm.ScaleMode = vbTwips
End Sub
Private Sub mnuWindow_Click()
On Error Resume Next
Dim bNothing As Boolean
bNothing = ActiveForm Is Nothing
bNothing = Not (bNothing)
mnuWindowArrangeIcons.Enabled = bNothing
mnuWindowCascade.Enabled = bNothing
mnuWindowMinimize.Enabled = bNothing
mnuWindowMinimizeAll.Enabled = bNothing
mnuWindowTileHorizontal.Enabled = bNothing
mnuWindowTileVertical.Enabled = bNothing
mnuWindowRestoreDown.Enabled = bNothing
mnuWindowRestoreUp.Enabled = bNothing
mnuWindowNext.Enabled = Forms.Count - 2 > 1
mnuWindowMinimize.Enabled = ActiveForm.WindowState <> 1
mnuWindowRestoreDown.Enabled = ActiveForm.WindowState <> 0
mnuWindowRestoreUp.Enabled = ActiveForm.WindowState <> 2
End Sub
Private Sub mnuFormatBullet_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SetFocus
SendKeys "^+L"
10:
End Sub



Private Sub mnuEditClear_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Local Error Resume Next
    If ActiveForm.rtfText.SelProtected = True Then Exit Sub
    StatusBar.Panels(1).Text = "Clearing..."
ActiveForm.rtfText.SelText = vbNullString
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub mnuEditCopyList_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
If ActiveForm.rtfText.SelText <> "" Then
frmCopylist.lstList.AddItem ActiveForm.rtfText.SelText
End If
10:
End Sub

Private Sub lblColor_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
Dim lngColor As Long
    lngColor = ShowColorDlg
    If lngColor <> -1 Then
        cFontColor(btFont) = lngColor
    End If
    lblColor.BackColor = cFontColor(btFont)
    ActiveForm.rtfText.SelColor = cFontColor(btFont)
ActiveForm.rtfText.SetFocus
10:
End Sub

Private Sub cboFontFace_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
If cboFontFace.Text = "" Then
    txtPreview.Visible = False
    Exit Sub
End If
vFontFace(btFont) = cboFontFace.Text
ActiveForm.rtfText.SelFontName = vFontFace(btFont)
ActiveForm.rtfText.SetFocus
txtPreview.Text = " " & ActiveForm.rtfText.SelFontName
10:
End Sub

Private Sub cboFontSize_Click()
On Local Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
cboFontSize.Text = Val(cboFontSize.Text)
If Val(cboFontSize.Text) > 1638.3 Then cboFontSize.Text = "1638.3"
If Val(cboFontSize.Text) < 1 Then cboFontSize.Text = "1"
intFontSize(btFont) = CInt(cboFontSize.Text)
ActiveForm.rtfText.SetFocus
ActiveForm.rtfText.SelFontSize = cboFontSize.Text
End Sub

Private Sub mnuFormatDecreaseHangingIndent_Click()
On Error Resume Next
ActiveForm.rtfText.SelHangingIndent = ActiveForm.rtfText.SelHangingIndent - 1
End Sub

Private Sub mnuViewExtendedFormatting_Click()
On Error Resume Next
    tbFormat.Visible = Not (tbFormat.Visible)
    mnuViewExtendedFormatting.Checked = tbFormat.Visible
End Sub

Public Sub ChangeFont(btFontIndex As Byte)
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
btFont = btFontIndex
cboFont.ListIndex = btFontIndex
ActiveForm.rtfText.SelFontName = vFontFace(btFontIndex)
ActiveForm.rtfText.SelFontSize = intFontSize(btFontIndex)
ActiveForm.rtfText.SelColor = cFontColor(btFontIndex)
cboFontSize.Text = intFontSize(btFontIndex)
cboFontFace.Text = vFontFace(btFontIndex)
ActiveForm.rtfText.SelColor = cFontColor(btFontIndex)
lblColor.BackColor = cFontColor(btFontIndex)
cboFontSize.Text = intFontSize(btFontIndex)
cboFontFace.Text = vFontFace(btFontIndex)
ActiveForm.rtfText.SelBold = FontBold(btFontIndex)
ActiveForm.rtfText.SelItalic = FontItalic(btFontIndex)
ActiveForm.rtfText.SelUnderline = FontUnderline(btFontIndex)
ActiveForm.rtfText.SelStrikeThru = FontStrikethru(btFontIndex)
cboFontFace_Click
ShowAttributes
10:
ErrorTrap "changing font"
End Sub

Private Sub mnuInsertDateTime_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
StatusBar.Panels(1).Text = "Inserting the Current Date & Time..."
ActiveForm.rtfText.SelText = DateTime.Date + DateTime.Time
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuFileImport_Click()
On Error GoTo 10
Dim sFile As String
If Not ActiveForm Is Nothing Then ActiveForm.rtfText.SetFocus
sFile = ShowCommonDlg(True, "", Me, "Text Files (*.rtf, *.wri, *.doc, *.text, *.txt)" & Chr(0) & "*.rtf;*.wri;*.doc;*.text;*.txt" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Insert File...", 4096)
If sFile = "" Then Exit Sub
If ActiveForm Is Nothing Then LoadNewDoc
StatusBar.Panels(1).Text = "Opening Document..."
On Error GoTo 10
StatusBar.Panels(1).Text = "Importing..."
ActiveForm.rtfText.SelRTF = OpenBinary(sFile)
StatusBar.Panels(1).Text = "Ready"
Exit Sub
10:
ErrorTrap "inserting a file"
End Sub
Private Function OpenBinary(sFile As String) As String
Dim FileNum As Long
FileNum = FreeFile()
Open sFile For Binary As #FileNum
OpenBinary = Input(LOF(FileNum), #FileNum)
Close #FileNum
End Function
Private Sub mnuFormatIncreaseHangingIndent_Click()
On Error Resume Next
ActiveForm.rtfText.SelHangingIndent = ActiveForm.rtfText.SelHangingIndent + 1
End Sub

Private Sub mnuInsertDate_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
StatusBar.Panels(1).Text = "Inserting Fixed Date..."
ActiveForm.rtfText.SelText = DateTime.Date
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuFormatCaseLowerCase_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SelText = LCase(ActiveForm.rtfText.SelText)
10:
End Sub

Private Sub MDIForm_Initialize()
On Error GoTo 10
InitCommonControls
10:
ErrorTrap
End Sub
Private Sub ShowAttributes()
On Error Resume Next
tbFormat.Buttons(1).Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
tbFormat.Buttons(2).Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
tbFormat.Buttons(3).Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
tbFormat.Buttons(4).Value = IIf(ActiveForm.rtfText.SelStrikeThru, tbrPressed, tbrUnpressed)
End Sub
Private Sub MDIForm_Load()
On Error GoTo 10
Dim i As Long, bLoadPrefs As Boolean
Dim sFile As String
    If KeyDown(vbKeyControl) = True Then
        If CustomBox("Do you want to reset the preferences file?", "If you choose Reset, the preferences you set will be lost.", _
            vbExclamation, "", "&Cancel", "&Reset") = 1 Then
            ResetPrefs
        End If
    End If
    mnuRightClick.Visible = False
    bNormal = True
    LoadNewDoc
    AddShortcuts mnuFileSaveAs, "Shift+Ctrl+S"
    AddShortcuts mnuInsertNonbreakingSpace, "Ctrl+Space"
    AddShortcuts mnuEditDelPrevWord, "Ctrl+Bksp"
    AddShortcuts mnuEditDelNextWord, "Ctrl+Del"
    AddShortcuts mnuGoToEnd, "Ctrl+End"
    AddShortcuts mnuEditGoTo, "Ctrl+Home"
    AddShortcuts mnuGoToLineAbove, "Ctrl+Up Arrow"
    AddShortcuts mnuGoToLineBelow, "Ctrl+Dn Arrow"
    AddShortcuts mnuInsertAccentGrave, "Ctrl+`"
    AddShortcuts mnuInsertAccentAcute, "Ctrl+'"
    AddShortcuts mnuInsertAccentTilde, "Ctrl+Shift+`"
    AddShortcuts mnuInsertAccentUmlaut, "Ctrl+;"
    AddShortcuts mnuInsertAccentCaret, "Ctrl+Shift+6"
    AddShortcuts mnuInsertAccentCedilla, "Ctrl+,"
    'AddShortcuts mnuFormatToggleSmartQuotes, "Ctrl+Shift+'"
    AddShortcuts mnuGoToPrevWord, "Ctrl+Left"
    AddShortcuts mnuGoToNextWord, "Ctrl+Right"
    AddShortcuts mnuEditSelectAll, "Ctrl+A"
    AddShortcuts mnuFormatCaseToggleCaps, "Ctrl+Shift+A"
    AddShortcuts mnuWindowNext, "Ctrl+F6"
    AddShortcuts mnuFormatAlignLeft, "Ctrl+L"
    AddShortcuts mnuFormatAlignCenter, "Ctrl+E"
    AddShortcuts mnuFormatAlignRight, "Ctrl+R"
    AddShortcuts mnuFormatAlignJustify, "Ctrl+J"
    AddShortcuts mnuFormatSuperscript, "Ctrl+Shift+="
    AddShortcuts mnuFormatSubscript, "Ctrl+="
    AddShortcuts mnuEditCopy, "Ctrl+C"
    AddShortcuts mnuEditPaste, "Ctrl+V"
    AddShortcuts mnuEditCut, "Ctrl+X"
    AddShortcuts mnuEditSelectNextWord, "Ctrl+Shift+,"
    AddShortcuts mnuEditSelectPrevWord, "Ctrl+Shift+."
    AddShortcuts mnuFormatBullet, "Ctrl+Shift+L"
    bLoadPrefs = True
    GetRecentFiles
    If DoPrefs(0, "SaveWorkspace") = 1 Then
        cbCoolBar.Visible = DoPrefs(0, "ShowCoolBar") = "1"
        Toolbar.Visible = DoPrefs(0, "ShowToolbar") = "1"
        cbCoolBar.Bands(1).Visible = DoPrefs(0, "ShowToolbar") = "1"
        tbrFontFormat.Visible = DoPrefs(0, "ShowFormatBar") = "1"
        cbCoolBar.Bands(2).Visible = DoPrefs(0, "ShowFormatBar") = "1"
        tbFormat.Visible = DoPrefs(0, "ShowExtFormatting") = "1"
        tbSymbols.Visible = DoPrefs(0, "ShowSymbolBar") = "1"
        StatusBar.Visible = DoPrefs(0, "ShowStatusBar") = "1"
        pcSlider.Visible = DoPrefs(0, "ShowRuler") = "1"
        cbCoolBar.Bands(3).Visible = DoPrefs(0, "ShowRuler") = "1"
        Me.WindowState = DoPrefs(0, "WindowState")
    End If
    bLoadPrefs = False
    StatusBar.Style = sbrSimple
    DoToolbars
    DoEvents
    StatusBar.SimpleText = "Getting Fonts..."
For i = 0 To Screen.FontCount - 1
    cboFontFace.AddItem Screen.Fonts(i)
Next
If Command <> "" Then
    sFile = Mid$(Command, 2, Len(Command) - 2)
    Select Case Right$(sFile, 4)
        Case ".opb"
            OpenBook sFile
        Case ".rtf", ".wri", ".doc"
            OpenFile sFile, True, 0
        Case Else
            OpenFile sFile, True, 1
    End Select
End If
For i = 0 To 4
    cFontColor(i) = &H0
Next
cboFontSize.ListIndex = 2
      vFontFace(1) = "Courier New"
      intFontSize(1) = 10
      cFontColor(1) = 0
      btFontIndex(1) = 1
      vFontFace(0) = "Times New Roman"
      intFontSize(0) = 12
      cFontColor(0) = 0
      btFontIndex(0) = 1
      vFontFace(2) = "Times New Roman"
      intFontSize(2) = 14
      cFontColor(2) = 0
      btFontIndex(2) = 1
      FontBold(2) = True
      FontItalic(2) = True
      vFontFace(3) = "Times New Roman"
      intFontSize(3) = 16
      cFontColor(3) = 0
      btFontIndex(3) = 1
      FontBold(3) = True
      vFontFace(4) = "Times New Roman"
      intFontSize(4) = 8
      cFontColor(4) = 0
      btFontIndex(4) = 1
      cboFont.ListIndex = 0
      ReDim CustomColors(0 To 16 * 4 - 1) As Byte
      For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
      Next i
StatusBar.Style = sbrNormal
StatusBar.Panels(1).Text = "Ready"
MousePointer = 0
Exit Sub
10:
If bLoadPrefs = True Then
    Select Case Err.Number
        Case 0
        Case Else
            CustomBox "An error occurred while loading preferences file.", "The preferences file is invalid. If this causes problems, restart Hyperwrite. If problems still occur, reset the preferences file by choosing Reset in the Preferences dialog.", vbCritical, "", "", "&OK"
    End Select
End If
ErrorTrap "loading"
End Sub
Private Sub DoToolbars(Optional bSendMessage As Boolean = False)
    Dim strIconPath As String
    Const TBSTYLE_FLAT As Long = &H800
    Const TB_SETSTYLE = WM_USER + 56
    Const TB_GETSTYLE = WM_USER + 57
    On Error GoTo 10
    StatusBar.SimpleText = "Loading Toolbars..."
    If bSendMessage = True Then 'GlyphLab Code
        ' Set up the toolbar
        Dim lngStyle As Long
        Dim lRes As Long
        
        ' Get the toolbar handle (we cannot just use tbrX.hwnd as this is a container
        ' window for the actual toolbar control)
        Dim hTBar As Long
        Dim fTBar As Long
        Dim sTBar As Long
        hTBar = FindWindowEx(Toolbar.hwnd, 0&, "ToolbarWindow32", vbNullString)
        fTBar = FindWindowEx(tbFormat.hwnd, 0&, "ToolbarWindow32", vbNullString)
        sTBar = FindWindowEx(tbSymbols.hwnd, 0&, "ToolbarWindow32", vbNullString)
        
        ' The style "TBSTYLE_FLAT" needs to be added.  Although this option is available
        ' in the property pages for the toolbar, it needs to be set here.
        
        ' Get the current style
        lngStyle = SendMessage(hTBar, TB_GETSTYLE, 0&, ByVal 0&)
        lngStyle = SendMessage(fTBar, TB_GETSTYLE, 0&, ByVal 0&)
        lngStyle = SendMessage(sTBar, TB_GETSTYLE, 0&, ByVal 0&)
        
        ' Add the TBSTYLE_FLAT style (could also apply other styles here)
        lngStyle = lngStyle Or TBSTYLE_FLAT
            
        ' Set the new style
        Call SendMessage(hTBar, TB_SETSTYLE, 0&, ByVal lngStyle)
        Call SendMessage(fTBar, TB_SETSTYLE, 0&, ByVal lngStyle)
        Call SendMessage(sTBar, TB_SETSTYLE, 0&, ByVal lngStyle)
        Toolbar.Refresh
        tbFormat.Refresh
        tbSymbols.Refresh
        Exit Sub
    End If
    
        ' Set path to icon directory
        strIconPath = App.Path & "\icons\"
    
    With ImageList.ListImages
        'Standard Toolbar
        .Add , "New", LoadPicture(strIconPath + "new.bmp")
        .Add , "Open", LoadPicture(strIconPath + "open.bmp")
        .Add , "Save", LoadPicture(strIconPath + "save.bmp")
        .Add , "Print", LoadPicture(strIconPath + "print.bmp")
        .Add , "Find", LoadPicture(strIconPath + "find.bmp")
        .Add , "Cut", LoadPicture(strIconPath + "cut.bmp")
        .Add , "Copy", LoadPicture(strIconPath + "copy.bmp")
        .Add , "Paste", LoadPicture(strIconPath + "paste.bmp")
        .Add , "Undo", LoadPicture(strIconPath + "undo.bmp")
        .Add , "Redo", LoadPicture(strIconPath + "redo.bmp")
        .Add , "Image", LoadPicture(strIconPath + "insimg.bmp")
        .Add , "DateTime", LoadPicture(strIconPath + "datetime.bmp")
        .Add , "Symbol", LoadPicture(strIconPath + "symbol.bmp")
        'Format Toolbar
        .Add , "Bold", LoadPicture(strIconPath + "bold.bmp")
        .Add , "Italic", LoadPicture(strIconPath + "italic.bmp")
        .Add , "Underline", LoadPicture(strIconPath + "underline.bmp")
        .Add , "Strikethru", LoadPicture(strIconPath + "strikethru.bmp")
        .Add , "Left", LoadPicture(strIconPath + "left.bmp")
        .Add , "Center", LoadPicture(strIconPath + "center.bmp")
        .Add , "Right", LoadPicture(strIconPath + "right.bmp")
        .Add , "Bullets", LoadPicture(strIconPath + "bullets.bmp")
        .Add , "Superscript", LoadPicture(strIconPath + "superscript.bmp")
        .Add , "Subscript", LoadPicture(strIconPath + "subscript.bmp")
        .Add , "Size Plus", LoadPicture(strIconPath + "sizeplus.bmp")
        .Add , "Size Minus", LoadPicture(strIconPath + "sizeminus.bmp")
    End With
    
    ' Bind toolbar to imagelist and set buttons on toolbar which require icons
    Toolbar.ImageList = ImageList
    tbFormat.ImageList = ImageList
    
    ' Set an icon for each button on the toolbar
    Toolbar.Buttons(1).Image = "New"
    Toolbar.Buttons(2).Image = "Open"
    Toolbar.Buttons(3).Image = "Save"
    Toolbar.Buttons(4).Image = "Print"
    Toolbar.Buttons(6).Image = "Find"
    Toolbar.Buttons(7).Image = "Cut"
    Toolbar.Buttons(8).Image = "Copy"
    Toolbar.Buttons(9).Image = "Paste"
    Toolbar.Buttons(10).Image = "Undo"
    Toolbar.Buttons(11).Image = "Redo"
    Toolbar.Buttons(13).Image = "Image"
    Toolbar.Buttons(14).Image = "DateTime"
    Toolbar.Buttons(15).Image = "Symbol"
    tbFormat.Buttons(1).Image = "Bold"
    tbFormat.Buttons(2).Image = "Italic"
    tbFormat.Buttons(3).Image = "Underline"
    tbFormat.Buttons(4).Image = "Strikethru"
    tbFormat.Buttons(6).Image = "Left"
    tbFormat.Buttons(7).Image = "Center"
    tbFormat.Buttons(8).Image = "Right"
    tbFormat.Buttons(10).Image = "Bullets"
    tbFormat.Buttons(12).Image = "Superscript"
    tbFormat.Buttons(13).Image = "Subscript"
    tbFormat.Buttons(14).Image = "Size Plus"
    tbFormat.Buttons(15).Image = "Size Minus"
    StatusBar.SimpleText = ""
    StatusBar.Style = sbrNormal
    Exit Sub
10:
    CustomBox "An Error Occurred While Loading Toolbar Icons.", "Make sure there are 23 icons in " & App.Path & "\icons.", vbCritical, "", "", "&OK"
End Sub
Private Sub AddShortcuts(mnuMenuItem As Menu, strShortcut As String)
mnuMenuItem.Caption = mnuMenuItem.Caption & Chr(9) & strShortcut
End Sub
Private Sub GetRecentFiles()
If DoPrefs(0, "RecentFiles") = 0 Then
    mnuFileOpenRecent.Enabled = False
    Exit Sub
End If
Dim i As Integer
For i = 0 To 4
    mnuFileRecent(i).Caption = "&" & i + 1 & " " & ParseFileName(DoPrefs(0, "Recent" & i + 1))
    If mnuFileRecent(i).Caption = "" Then
    mnuFileRecent(i).Visible = False
    End If
Next
End Sub
Private Sub SaveRecentFiles(FileName As String)
If DoPrefs(0, "RecentFiles") = 0 Then
    DoPrefs 1, "Recent1", ""
    DoPrefs 1, "Recent2", ""
    DoPrefs 1, "Recent3", ""
    DoPrefs 1, "Recent4", ""
    DoPrefs 1, "Recent5", ""
    Exit Sub
End If
    Dim RegFileName(4) As String
        RegFileName(0) = DoPrefs(0, "Recent1")
        RegFileName(1) = DoPrefs(0, "Recent2")
        RegFileName(2) = DoPrefs(0, "Recent3")
        RegFileName(3) = DoPrefs(0, "Recent4")
        RegFileName(4) = DoPrefs(0, "Recent5")
        If RegFileName(4) = "" Or RegFileName(3) = "" Or RegFileName(2) = "" _
        Or RegFileName(1) = "" Or RegFileName(0) = "" Then
          DoPrefs 1, "Recent5", RegFileName(3)
          DoPrefs 1, "Recent4", RegFileName(2)
          DoPrefs 1, "Recent3", RegFileName(1)
          DoPrefs 1, "Recent2", RegFileName(0)
          DoPrefs 1, "Recent1", FileName
        End If
        If FileName = RegFileName(0) Or FileName = RegFileName(1) Or FileName = RegFileName(2) _
        Or FileName = RegFileName(3) Or FileName = RegFileName(4) Then Exit Sub
        If FileName <> RegFileName(0) Then
            DoPrefs 1, "Recent1", FileName
            DoPrefs 1, "Recent2", RegFileName(0)
            DoPrefs 1, "Recent3", RegFileName(1)
            DoPrefs 1, "Recent4", RegFileName(2)
            DoPrefs 1, "Recent5", RegFileName(3)
        End If
7:
End Sub
Private Sub mnuWindowRestoreDown_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.WindowState = 0
10:
End Sub

Private Sub mnuEditGoTo_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelStart = 0
10:
End Sub

Private Sub mnuEdit_Click()
On Error GoTo 10
DoMenus
If ActiveForm Is Nothing Then Exit Sub
    Dim TextSelected As Boolean
    TextSelected = ActiveForm.rtfText.SelLength > 0
    mnuEditCut.Enabled = TextSelected
    mnuEditCopy.Enabled = TextSelected
    mnuEditClear.Enabled = TextSelected
    mnuEditCopyList.Enabled = TextSelected
    mnuEditSelectNone.Enabled = TextSelected
    mnuEditTrimSpaces.Enabled = TextSelected
    mnuEditDelNextWord.Enabled = TextSelected
    mnuEditDelPrevWord.Enabled = TextSelected
    If mnuEditUndo.Enabled = True Then
        mnuEditUndo.Caption = "&Undo " & TranslateUndoType(UndoType) & vbTab & "Ctrl+Z"
    End If
    If TextSelected = True Then
        mnuEditChgProtection.Caption = "Change Protection for Selection"
    Else
        mnuEditChgProtection.Caption = "Change Protection for Document"
    End If
    If ActiveForm Is Nothing Then
    Else
        mnuEditUndoReplace.Enabled = ActiveForm.rtfText.Tag <> ""
    End If
10:
End Sub

Private Function ShowColorDlg() As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    cc.lStructSize = Len(cc)
    cc.hwndOwner = Me.hwnd
    cc.hInstance = 0
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    lReturn = ChooseColorAPI(cc)
    If lReturn <> 0 Then
         ShowColorDlg = cc.rgbResult
    Else
         ShowColorDlg = -1
     End If
    CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
End Function

Private Sub LoadNewDoc()
On Error GoTo 10
    If btDocumentCount > 255 Then
        CustomBox "There are too many windows open.", "Close a few windows and try again to open a new one.", vbExclamation, "", "", "&OK"
        Exit Sub
    End If
    btDocumentCount = btDocumentCount + 1
    Set frmDocument = New frmDocument
    frmDocument.Caption = ""
    frmDocument.Show
    If btDocumentCount = 1 Then
        ActiveForm.Caption = "untitled"
    Else
        ActiveForm.Caption = "untitled " & btDocumentCount
    End If
    Slider.SelStart = 0
10:
End Sub

Private Sub mnuEditRedo_Click()
Const EM_REDO = (&H400 + 84)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
SendMessage ActiveForm.rtfText.hwnd, EM_REDO, 0, 0
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
End Sub

Private Sub mnuFileAutoSave_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
If ActiveForm.rtfText.FileName = "" Then
    mnuFileSaveAs_Click
    If ActiveForm.rtfText.FileName = "" Then
        ActiveForm.bAutoSave = False
        mnuFileAutoSave.Checked = ActiveForm.bAutoSave
    End If
    Exit Sub
End If
ActiveForm.bAutoSave = Not (ActiveForm.bAutoSave)
mnuFileAutoSave.Checked = ActiveForm.bAutoSave
If ActiveForm.bAutoSave = True Then
    Dim intDuration As Integer
    intDuration = InputBox("Amount of changes between save:", "AutoSave")
    ActiveForm.lSaveStart = intDuration
    ActiveForm.CurrentStart = 0
End If
End Sub

Private Sub mnuFileCloseAll_Click()
On Error GoTo 10
Dim i As Integer
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
StatusBar.Panels(1).Text = "Closing All Windows..."
For i = 1 To Forms.Count - 2
Unload ActiveForm
Next
StatusBar.Panels(1).Text = "Ready"
10:
End Sub


Private Sub mnuFileRevert_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
If ActiveForm.rtfText.FileName = "" Then
CustomBox "No file loaded.", "Cannot revert to saved document because the document has not been saved yet.", vbExclamation, "", "", "OK"
Exit Sub
End If
On Error GoTo 10
Dim iMsgBoxReturn As Integer
iMsgBoxReturn = CustomBox("Are you sure you want to revert the document to the saved version?", "This will erase all changes made to the document since the last time you saved it.", vbExclamation, "", "Cancel", "Revert")
If iMsgBoxReturn = 1 Then
ActiveForm.rtfText.Text = ""
ActiveForm.rtfText.LoadFile ActiveForm.rtfText.FileName
End If
10:
End Sub

Private Sub mnuFileSaveAll_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim i As Integer
On Error GoTo 10
    For i = 1 To Forms.Count - 1
        mnuFileSave_Click
        SendKeys "^{F6}"
        DoEvents
    Next
StatusBar.Panels(1).Text = "Ready"
Exit Sub
10:
If ActiveForm Is Nothing Then
ErrorTrap "attempting to save all files"
Else
ErrorTrap "attempting to save " & ActiveForm.rtfText.FileName
End If
End Sub

Private Sub mnuFormatBold_Click()
On Error Resume Next
If ActiveForm.rtfText.SelBold = False Then
ActiveForm.rtfText.SelBold = True
Else
ActiveForm.rtfText.SelBold = False
End If
If ActiveForm.rtfText.SelBold = True Then
mnuFormatBold.Checked = True
Else
mnuFormatBold.Checked = False
End If
FontBold(btFont) = Not (FontBold(btFont))
End Sub
Private Sub mnuFormatdindent_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.ScaleMode = vbMillimeters
If ActiveForm.rtfText.SelIndent > 0 Then
ActiveForm.rtfText.SelIndent = ActiveForm.rtfText.SelIndent - 15
Slider.Value = ActiveForm.rtfText.SelIndent / 1.58
Else
ActiveForm.rtfText.SelIndent = 0
End If
ActiveForm.ScaleMode = vbTwips
End Sub

Private Sub mnuFormatiindent_Click()
On Error Resume Next
ActiveForm.ScaleMode = vbInches
If ActiveForm.rtfText.SelIndent > 0 Then
ActiveForm.rtfText.SelIndent = ActiveForm.rtfText.SelIndent + 0.5
Slider.Value = ActiveForm.rtfText.SelIndent * 16
Else
ActiveForm.rtfText.SelIndent = 1
End If
ActiveForm.ScaleMode = vbTwips
End Sub

Private Sub mnuFormatItalic_Click()
On Error Resume Next
If ActiveForm.rtfText.SelItalic = False Then
ActiveForm.rtfText.SelItalic = True
Else
ActiveForm.rtfText.SelItalic = False
End If
If ActiveForm.rtfText.SelItalic = True Then
mnuFormatItalic.Checked = True
Else
mnuFormatItalic.Checked = False
End If
FontItalic(btFont) = Not (FontItalic(btFont))
10:
End Sub

Private Sub mnuFormatstrikethru_Click()
On Error Resume Next
If ActiveForm.rtfText.SelStrikeThru = False Then
ActiveForm.rtfText.SelStrikeThru = True
Else
ActiveForm.rtfText.SelStrikeThru = False
End If
If ActiveForm.rtfText.SelStrikeThru = True Then
mnuFormatstrikethru.Checked = True
Else
mnuFormatstrikethru.Checked = False
End If
FontStrikethru(btFont) = Not (FontStrikethru(btFont))
10:
End Sub

Private Sub mnuFormatUnderline_Click()
On Error Resume Next
If ActiveForm.rtfText.SelUnderline = False Then
ActiveForm.rtfText.SelUnderline = True
Else
ActiveForm.rtfText.SelUnderline = False
End If
If ActiveForm.rtfText.SelUnderline = True Then
mnuFormatUnderline.Checked = True
Else
mnuFormatUnderline.Checked = False
End If
FontUnderline(btFont) = Not (FontUnderline(btFont))
10:
End Sub

Private Sub mnuInsertCharacter_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmCharacters.Show vbModal, Me
10:
ErrorTrap "attempting to show Special Character dialog"
End Sub

Private Sub mnuViewAccentsBar_Click()
    tbSymbols.Visible = Not (tbSymbols.Visible)
    mnuViewAccentsBar.Checked = tbSymbols.Visible
End Sub

Private Sub mnuWindowMinimize_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.WindowState = 1
10:
End Sub

Private Sub mnuWindowMinimizeAll_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim i As Integer
For i = 1 To Forms.Count - 1
 ActiveForm.WindowState = 1
Next
10:
End Sub

Private Sub mnuWindowNext_Click()
ActiveForm.rtfText.SetFocus
SendKeys "^{F6}"
End Sub

Private Sub mnuEditPasteList_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmCopylist.Show vbModal, Me
10:
End Sub

Private Sub mnuFileRecent_Click(Index As Integer)
Dim strReg As String
strReg = DoPrefs(0, "Recent" & Index + 1)
OpenFile strReg, False, , True
End Sub

Private Sub mnuFormatResetHangingIndent_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error Resume Next
ActiveForm.rtfText.SelHangingIndent = 0
End Sub

Private Sub mnuWindowRestoreUp_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.WindowState = 2
End Sub

Private Sub mnuViewRuler_Click()
On Error GoTo 10
    pcSlider.Visible = Not (pcSlider.Visible)
    cbCoolBar.Bands(3).Visible = pcSlider.Visible
    mnuViewRuler.Checked = pcSlider.Visible
10:
End Sub

Private Sub mnuEditSelectAll_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
StatusBar.Panels(1).Text = "Selecting all..."
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText.Text)
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuFileSaveSelection_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
If ActiveForm.rtfText.SelText <> "" Then
Dim sFile As String
    sFile = ShowCommonDlg(False, "rtf", Me, "Text Files (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr(0) & "*.rtf;*.wri;*.doc;*.txt;*.text" & Chr(0) & "Web Source Code (*.htm, *.html, *.xml, *.asp *.aspx, *.shtml, *.shtm, *.stm, *.php, *.css)" & Chr(0) & "*.htm;*.html;*.xml;*.asp;*.aspx;*.shtml;*.shtm;*.stm;*.stm;*.php;*.css" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Save As", 0)
    Dim FileNum%
    FileNum% = FreeFile
    Open sFile For Output As FileNum%
    Print #FileNum%, ActiveForm.rtfText.SelRTF
    Close #FileNum%
End If
10:
ErrorTrap "while saving current selection"
End Sub

Private Sub mnuInsertJavaScript_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelText = "<script type=" + sQuote + "text/javascript" + sQuote + ">" + vbNewLine + "<!--Your script goes here.-->" + vbNewLine + "</script>"
End Sub

Private Sub mnuInsertTime_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
StatusBar.Panels(1).Text = "Inserting current Time..."
ActiveForm.rtfText.SelText = DateTime.Time
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub statusbar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar.Panels(1).Text = "Ready"
End Sub


Private Sub mnuEditSelectNone_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
StatusBar.Panels(1).Text = "Selecting None..."
ActiveForm.rtfText.SelLength = 0
StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuInsertStartingXML_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelText = "<?xml version=" + sQuote + "1.0" + sQuote + " encoding=" + sQuote + "utf-8" + sQuote + "?>"
End Sub

Private Sub mnuFormatSubscript_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
'ActiveForm.rtfText.SelCharOffset = -55
ActiveForm.rtfText.SetFocus
SendKeys "^="
10:
End Sub

Private Sub mnuFormatSuperscript_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
'ActiveForm.rtfText.SElcharoffset = 55
ActiveForm.rtfText.SetFocus
SendKeys "^+="
10:
End Sub

Private Sub mnuHelpAbout_Click()
On Error Resume Next
StatusBar.Panels(1).Text = "Opening About Dialog..."
    StatusBar.Panels(1).Text = "Ready"
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuWindowArrangeIcons_Click()
On Error GoTo 10
StatusBar.Panels(1).Text = "Arranging Icons..."
    Me.Arrange vbArrangeIcons
    StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuWindowTileVertical_Click()
On Error GoTo 10
StatusBar.Panels(1).Text = "Tiling Windows..."
    Me.Arrange vbTileVertical
    StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuWindowTileHorizontal_Click()
On Error GoTo 10
StatusBar.Panels(1).Text = "Tiling Windows..."
    Me.Arrange vbTileHorizontal
    StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuWindowCascade_Click()
On Error GoTo 10
StatusBar.Panels(1).Text = "Cascading Current Windows..."
    Me.Arrange vbCascade
    StatusBar.Panels(1).Text = "Ready"
10:
End Sub

Private Sub mnuViewStatusBar_Click()
On Error GoTo 10
    StatusBar.Visible = Not (StatusBar.Visible)
    mnuViewStatusBar.Checked = StatusBar.Visible
10:
End Sub

Private Sub mnuViewToolbar_Click()
    Toolbar.Visible = Not (Toolbar.Visible)
    cbCoolBar.Bands(1).Visible = Toolbar.Visible
    mnuViewToolbar.Checked = Toolbar.Visible
End Sub

Private Sub mnuEditPaste_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
StatusBar.Panels(1).Text = "Pasting..."
SendMessage ActiveForm.rtfText.hwnd, WM_PASTE, 0, 0
StatusBar.Panels(1).Text = "Ready"
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
10:
End Sub

Private Sub mnuEditCopy_Click()
If ActiveForm Is Nothing Then HandleNoWindows
On Error GoTo 15
StatusBar.Panels(1).Text = "Copying..."
    SendMessage ActiveForm.rtfText.hwnd, WM_COPY, 0&, ByVal 0&
    StatusBar.Panels(1).Text = "Ready"
15:
End Sub

Private Sub mnuEditCut_Click()
If ActiveForm Is Nothing Then HandleNoWindows
On Error GoTo 15
StatusBar.Panels(1).Text = "Cutting..."
    SendMessage ActiveForm.rtfText.hwnd, WM_CUT, 0&, ByVal 0&
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
    StatusBar.Panels(1).Text = "Ready"
15:
End Sub

Private Sub mnuEditUndo_Click()
Const EM_UNDO = &HC7
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
StatusBar.Panels(1).Text = "Undoing previous action..."
'ActiveForm.Undo
    SendMessage ActiveForm.rtfText.hwnd, EM_UNDO, 0, 0
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
    StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click() 'Microsoft code
On Error GoTo 10
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
Dim printDlg As PrinterDlg
Set printDlg = New PrinterDlg
' Set the starting information for the dialog box based on the current
' printer settings.
printDlg.PrinterName = Printer.DeviceName
printDlg.DriverName = Printer.DriverName
printDlg.Port = Printer.Port

' Set the default PaperBin so that a valid value is returned even
' in the Cancel case.
printDlg.PaperBin = Printer.PaperBin

' Set the flags for the PrinterDlg object using the same flags as in the
' common dialog control. The structure starts with VBPrinterConstants.
printDlg.flags = VBPrinterConstants.cdlPDNoSelection _
                 Or VBPrinterConstants.cdlPDNoPageNums _
                 Or VBPrinterConstants.cdlPDReturnDC
Printer.TrackDefault = False
If Not printDlg.ShowPrinter(Me.hwnd) Then Exit Sub
ActiveForm.rtfText.SelPrint printDlg.hDC
10:
If Err.Number = 429 Then
    If CustomBox("An error occurred while attempting to print. Do you want to try and activate the print dialog library?", "The print dialog library could not be initialized. This can happen if it hasn't been registered or installed before. Registering it may correct this problem.", _
    vbCritical, "", "&Cancel", "&Register") = 1 Then
        Shell "regsvr32 " & Environ("WINDIR") & "\system32\VBPrnDlg.dll"
    End If
    Exit Sub
End If
If Err.Number = -2147467259 Then
    If CustomBox("Could not print because the Print Spooler service is not running and/or the startup type is incorrect.", "Error Number: -2147467259. Make sure the Print Spooler service is running and that its startup type is Automatic.", _
    vbCritical, "", "&Cancel", "&Activate") = 1 Then
        Shell "sc start Spooler"
    End If
    Exit Sub
End If
ErrorTrap "printing"
End Sub

Public Sub mnuFileSaveAs_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim iMsgBoxReturn As Integer
Dim sFile As String
sFile = ShowCommonDlg(False, "", Me, "Text Files (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr(0) & "*.rtf;*.wri;*.doc;*.txt;*.text" & Chr(0) & "Web Source Code (*.htm, *.html, *.xml, *.asp *.aspx, *.shtml, *.shtm, *.stm, *.php, *.css)" & Chr(0) & "*.htm;*.html;*.xml;*.asp;*.aspx;*.shtml;*.shtm;*.stm;*.stm;*.php;*.css" & Chr(0) & "All Files (*)" & Chr(0) & "*" & Chr(0), "Save As", 4096)
If sFile = "" Then Exit Sub
    If InStr("|.rtf|.wri|.doc", Right(sFile, 4)) <> 0 Then
        StatusBar.Style = sbrSimple
        StatusBar.SimpleText = "Saving..."
        ActiveForm.rtfText.SaveFile sFile, rtfRTF
        StatusBar.Style = sbrNormal
        StatusBar.SimpleText = ""
    Else
        If DoPrefs(0, "WarnTextFormat") = "1" Then
            iMsgBoxReturn = CustomBox("Saving to this format will cause all existing formatting to be lost.", "This format doesn" & sApostrophe & "t support formatting that you may have put in your document.", vbExclamation, "", "Don" & sApostrophe & "t Save", "Save")
            If iMsgBoxReturn = 0 Then Exit Sub
        End If
        StatusBar.Style = sbrSimple
        StatusBar.SimpleText = "Saving..."
        ActiveForm.rtfText.SaveFile sFile, rtfText
        StatusBar.Style = sbrNormal
        StatusBar.SimpleText = ""
    End If
    ActiveForm.rtfText.FileName = sFile
    ActiveForm.Caption = ParseFileName(ActiveForm.rtfText.FileName)
    fMainForm.mnuFileSave.Enabled = False
    SaveRecentFiles (ActiveForm.rtfText.FileName)
    GetRecentFiles
10:
If Err.Number = 75 Then
    If CustomBox("This file cannot be saved due to an access error.", "The file might be read-only or another application may be using it. Would you like to save the file to another location?", vbExclamation, "", "&Cancel", "&Save As...") = 1 Then mnuFileSaveAs_Click
    Exit Sub
End If
ErrorTrap "saving " & ParseFileName(sFile) & " as a new file", sFile
End Sub

Public Sub mnuFileSave_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
Exit Sub
End If
On Error GoTo 10
Dim intMsgReturn As Integer
Dim strExtension As String
If ActiveForm.rtfText.FileName = "" Then
    mnuFileSaveAs_Click
    Exit Sub
End If
StatusBar.Panels(1).Text = "Saving..."
strExtension = Right$(ActiveForm.rtfText.FileName, 4)
If InStr("|.rtf|.wri|.doc", strExtension) <> 0 Then
    ActiveForm.rtfText.SaveFile ActiveForm.rtfText.FileName, rtfRTF
Else
    If DoPrefs(0, "WarnTextFormat") = "1" Then
        intMsgReturn = CustomBox("Saving to this format will cause all existing formatting to be lost.", "This format doesn" & sApostrophe & "t support formatting that you may have put in your document.", vbExclamation, "", "Don" & sApostrophe & "t Save", "Save")
        If intMsgReturn = 1 Then
            ActiveForm.rtfText.SaveFile ActiveForm.rtfText.FileName, rtfText
        Else
            Exit Sub
        End If
    End If
End If
StatusBar.Panels(1).Text = "Ready"
ActiveForm.bChanged = False
fMainForm.mnuFileSave.Enabled = False
ActiveForm.Caption = ParseFileName(ActiveForm.rtfText.FileName)
ActiveForm.rtfText.SetFocus
If ActiveForm.bAutoSave = True Then
    ActiveForm.CurrentStart = ActiveForm.rtfText.SelStart
End If
Exit Sub
10:
If Err.Number = 75 Then
    If CustomBox("This file cannot be saved due to an access error.", "The file might be read-only or another application may be using it. Would you like to save the file to another location?", vbExclamation, "", "&Cancel", "&Save As...") = 1 Then mnuFileSaveAs_Click
    Exit Sub
End If
ErrorTrap ActiveForm.rtfText.FileName
End Sub

Private Sub mnuFileClose_Click()
If ActiveForm Is Nothing Then GoTo 10
StatusBar.Panels(1).Text = "Closing..."
Unload ActiveForm
StatusBar.Panels(1).Text = "Ready"
Exit Sub
10:
HandleNoWindows
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo 10
Dim sFile As String
sFile = ShowCommonDlg(True, "rtf", Me, _
"Text Files (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr(0) & _
"*.rtf;*.wri;*.doc;*.txt;*.text" & Chr(0) & _
"Web Source Code (*.htm, *.html, *.xml, *.asp, *.aspx, *.shtml, *.shtm, *.stm, *.php, *.css)" _
& Chr(0) & "*.htm;*.html;*.xml;*.asp;*.aspx;*.shtml;*.shtm;*.stm;*.stm;*.php;*.css" & Chr(0) _
& "All Files (*)" & Chr(0) & "*" & Chr(0), "Open", 4096)
If sFile = "" Then Exit Sub
'sFile = Left(sFile, InStr(1, sFile, vbNullChar) - 1)
Dim sFileNameExt As String
sFileNameExt = Right$(sFile, 4)
OpenFile sFile, , , True
If DoPrefs(0, "AutoReplaceStraightQuotes") = "1" Then
    mnuFormatReplaceDQ_Click
End If
10:
ErrorTrap "opening " & sFile, sFile
End Sub
Private Sub OpenFile(sFile As String, Optional bRecent As Boolean = True, _
    Optional btOpenConst As Byte = 0, Optional bExplicitDetect As Boolean = False)
On Error GoTo 10
Dim sFileNameExt As String
    sFileNameExt = Right$(sFile, 4)
    If ActiveForm Is Nothing Then
        LoadNewDoc
    Else
        If ActiveForm.rtfText.Text <> "" Then LoadNewDoc
    End If
    If DoPrefs(0, "ImportPictures") = "1" And btOpenConst <> rtfText Then
        If InStr("|.jpg|.jpeg|.gif|.bmp|.dib|.wmf|.emf", sFileNameExt) <> 0 Then
            InsertObj (sFile)
            ActiveForm.rtfText.FileName = ""
            Exit Sub
        End If
    End If
    If bExplicitDetect = True Then
        If InStr("|.rtf|.wri|.doc", Right(sFile, 3)) <> 0 Then
            btOpenConst = 0
        Else
            btOpenConst = 1
        End If
    End If
    ActiveForm.rtfText.LoadFile sFile, btOpenConst
    Slider.Value = ActiveForm.rtfText.SelIndent / 90 'Make the slider show the left indent for the document
    ActiveForm.Caption = ParseFileName(sFile)
    If bRecent = True Then
        SaveRecentFiles (ActiveForm.rtfText.FileName)
        GetRecentFiles
    End If
    StatusBar.Panels(1).Text = "Ready"
    ActiveForm.bChanged = False
10:
ErrorTrap "opening " & sFile, sFile
End Sub

Private Sub mnuFileNew_Click()
StatusBar.Panels(1).Text = "Creating New Document..."
    LoadNewDoc
    StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub Slider_Scroll()
On Error GoTo 10
ActiveForm.ScaleMode = vbInches
ActiveForm.rtfText.SelIndent = Slider.Value / 16
ActiveForm.rtfText.SelRightIndent = Slider.Value / 16
ActiveForm.ScaleMode = vbTwips
10:
End Sub

Private Sub tbFormat_ButtonClick(ByVal Button As ComctlLib.Button)
On Error GoTo 20
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
    Select Case Button.Key
    Case "Bold"
        ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
        Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        FontBold(btFont) = Not (FontBold(btFont))
    Case "Italic"
        ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
        Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        FontItalic(btFont) = Not (FontItalic(btFont))
    Case "Underline"
        ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
        Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        FontUnderline(btFont) = Not (FontUnderline(btFont))
    Case "Strikethru"
        ActiveForm.rtfText.SelStrikeThru = Not ActiveForm.rtfText.SelStrikeThru
        Button.Value = IIf(ActiveForm.rtfText.SelStrikeThru, tbrPressed, tbrUnpressed)
        FontStrikethru(btFont) = Not (FontStrikethru(btFont))
    Case "Left"
        ActiveForm.rtfText.SelAlignment = rtfLeft
    Case "Center"
        ActiveForm.rtfText.SelAlignment = rtfCenter
    Case "Right"
        ActiveForm.rtfText.SelAlignment = rtfRight
    Case "Bullets"
        mnuFormatBullet_Click
    Case "Superscript"
        mnuFormatSuperscript_Click
    Case "Subscript"
        mnuFormatSubscript_Click
    Case "Increase Size"
        mnuFormatChgSize_Click (0)
    Case "Decrease Size"
        mnuFormatChgSize_Click (1)
    End Select

20:
End Sub

Private Sub tbsymbols_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SelText = Button.Caption
End Sub

Private Sub tbsymbols_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Static bShifted As Boolean
Dim i As Integer
If Button = vbRightButton Then bShifted = Not (bShifted)
If bShifted = True Then
For i = 16 To 39
tbSymbols.Buttons(i).Caption = UCase(tbSymbols.Buttons(i).Caption)
Next
Else
For i = 16 To 39
tbSymbols.Buttons(i).Caption = LCase(tbSymbols.Buttons(i).Caption)
Next
End If
End Sub

Private Sub tbsymbols_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bLiveWC = False Then
    StatusBar.Panels(1).Text = "Right click to toggle case"
Else
    FlashStatus "Right click to toggle case", 3
End If
End Sub

Private Sub tmrLiveWC_Timer()
If Not ActiveForm Is Nothing Then
    StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
End If
tmrLiveWC.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
StatusBar.SimpleText = "Ready"
StatusBar.Style = sbrNormal
If bLiveWC = False Then
    StatusBar.Panels(1).Text = "Ready"
Else
    tmrLiveWC_Timer
End If
tmrTimer.Enabled = False
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Find/Change"
            mnuEditfindReplace_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Undo"
            mnuEditUndo_Click
        Case "Redo"
            mnuEditRedo_Click
        Case "Insert Image"
            mnuInsertObject_Click
        Case "Insert Date and Time"
            mnuInsertDateTime_Click
        Case "Insert Symbol"
            mnuInsertCharacter_Click
    End Select
End Sub

Private Sub txtFind_Change()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
bRealSymbols = False
If chkOptions(2).Value = Checked Then
FindFunc txtFind.Text, ActiveForm.rtfText, False
End If
If txtFind.Text = "" Then
    cmdFindNext.Enabled = False
    cmdFindPrev.Enabled = False
Else
    cmdFindNext.Enabled = True
    cmdFindPrev.Enabled = True
End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdFindNext_Click
bRealSymbols = False
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
bRealSymbols = False
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
bRealSymbols = False
End Sub

Private Sub txtPreview_Change()
If txtPreview.Text = "" Then Exit Sub
txtPreview.Font = ActiveForm.rtfText.SelFontName
txtPreview.Width = 1905
txtPreview.FontBold = False
txtPreview.FontItalic = False
txtPreview.Visible = True
End Sub

Private Sub txtPreview_Click()
cboFontFace.SelStart = txtPreview.SelStart
cboFontFace.SelLength = txtPreview.SelLength
txtPreview.Visible = False
End Sub

Private Sub txtReplace_Change()
cmdReplace.Caption = "&Replace"
End Sub

Private Sub txtreplace_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdReplace_Click
End Sub

Private Sub mnuViewToggleSlider_Click()
On Error GoTo 10
Slider.Enabled = Not (Slider.Enabled)
If Slider.Enabled = True Then
    mnuViewToggleSlider.Caption = "&Lock Slider"
Else
    mnuViewToggleSlider.Caption = "&Unlock Slider"
End If
10:
End Sub

Private Sub mnuEditTrimSpaces_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelText = Trim$(ActiveForm.rtfText.SelText)
If bLiveWC = True Then StatusBar.Panels(1).Text = WordCount(ActiveForm.rtfText.Text) & " words"
End Sub

Private Sub mnuFormatCaseUppercase_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SelText = UCase(ActiveForm.rtfText.SelText)
10:
End Sub

Private Sub mnuToolsDocStatistics_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
StatusBar.Panels(1).Text = "Generating Statistics..." & "Press ESC to cancel word count"
Dim Count As Long
If ActiveForm.rtfText.SelText = "" Then
Count = WordCount(ActiveForm.rtfText.Text)
StatusBar.Panels(1).Text = "Ready"
CustomBox "Document Report:", "There are " & Count & " words (" & Len(ActiveForm.rtfText.Text) & " characters, " & FindOccurrences(ActiveForm.rtfText.TextRTF, "\par") & " paragraphs, " & GetLineCount(ActiveForm.rtfText.hwnd) & " lines, around " & Format(Len(ActiveForm.rtfText.Text) / 3750, "0.0") & " Letter-sized pages worth of content) in your document.", vbInformation, "", "", "OK"
Else
Count = WordCount(ActiveForm.rtfText.SelText)
StatusBar.Panels(1).Text = "Ready"
CustomBox "Document Report:", "There are " & Count & " words (" & Len(ActiveForm.rtfText.SelText) & " characters, " & FindOccurrences(ActiveForm.rtfText.SelRTF, "\par") & " paragraphs, around " & CSng(Len(ActiveForm.rtfText.SelText)) / 3750 & " Letter-sized pages worth of content) in the selected text.", vbInformation, "", "", "OK"
End If
10:
End Sub
