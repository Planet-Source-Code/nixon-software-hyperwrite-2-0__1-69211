VERSION 5.00
Begin VB.Form frmGetInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info"
   ClientHeight    =   3180
   ClientLeft      =   2310
   ClientTop       =   2100
   ClientWidth     =   6885
   Icon            =   "frmFileOperations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFileName 
      BorderStyle     =   0  'None
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
      Left            =   840
      MousePointer    =   3  'I-Beam
      TabIndex        =   20
      Tag             =   "0"
      Text            =   "No File"
      Top             =   120
      Width           =   5865
   End
   Begin VB.CheckBox chkTemporary 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Temporary"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   2340
      Width           =   1230
   End
   Begin VB.CheckBox chkSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System"
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
      Left            =   1905
      TabIndex        =   8
      Top             =   2565
      Width           =   1065
   End
   Begin VB.CheckBox chkArchive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Archive"
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
      Left            =   825
      TabIndex        =   7
      Top             =   2565
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "D&one"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   2760
      Width           =   960
   End
   Begin VB.CheckBox chkHidden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hidden"
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
      Left            =   1905
      TabIndex        =   5
      Top             =   2340
      Width           =   1065
   End
   Begin VB.CheckBox chkReadOnly 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Read-Only"
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
      Left            =   825
      TabIndex        =   4
      Top             =   2340
      Width           =   1065
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "&Move…"
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
      Left            =   780
      TabIndex        =   3
      Top             =   1560
      Width           =   885
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy…"
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
      Left            =   1710
      TabIndex        =   2
      Top             =   1560
      Width           =   885
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   2700
      TabIndex        =   1
      Top             =   1560
      Width           =   795
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load…"
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
      Left            =   5910
      TabIndex        =   0
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lblKind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kind:"
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
      Left            =   825
      TabIndex        =   21
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblWhere 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Where:"
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
      Left            =   825
      TabIndex        =   19
      Top             =   1020
      Width           =   705
   End
   Begin VB.Label lblLocation 
      BackColor       =   &H00FFFFFF&
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   18
      Tag             =   "\"
      Top             =   1020
      Width           =   5100
   End
   Begin VB.Label lblKind1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nonexistent File"
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
      Left            =   1605
      TabIndex        =   17
      Top             =   480
      Width           =   5100
   End
   Begin VB.Label lblFormat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   570
      Width           =   525
   End
   Begin VB.Label lblAttr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No file loaded"
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
      Left            =   1740
      TabIndex        =   15
      Top             =   1980
      Width           =   4965
   End
   Begin VB.Label lblModCreate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modified:"
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
      Left            =   825
      TabIndex        =   14
      Top             =   1290
      Width           =   705
   End
   Begin VB.Label lblModified 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00/00/00 00:00:00 PM"
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
      Left            =   1605
      TabIndex        =   13
      Top             =   1290
      Width           =   5100
   End
   Begin VB.Label lblAttributes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attributes:"
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
      Left            =   825
      TabIndex        =   12
      Top             =   1980
      Width           =   825
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0 Bytes"
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
      Left            =   1605
      TabIndex        =   11
      Top             =   750
      Width           =   5100
   End
   Begin VB.Label lblFileSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Size:"
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
      Left            =   825
      TabIndex        =   10
      Top             =   750
      Width           =   705
   End
   Begin VB.Image imgFile 
      Height          =   705
      Left            =   150
      Picture         =   "frmFileOperations.frx":030A
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmGetInfo"
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
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Const vbTemporary = &H100
Private Const vbCompressed = &H800
Private FileNameStr As String


Private Sub ApplyAttributes()
On Error GoTo 10
Dim attr As Long
    If chkReadOnly.Value = Checked Then attr = vbReadOnly
    If chkArchive.Value = Checked Then attr = attr + vbArchive
    If chkSystem.Value = Checked Then attr = attr + vbSystem
    If chkHidden.Value = Checked Then attr = attr + vbHidden
    If chkTemporary.Value = Checked Then attr = attr + vbTemporary
    SetFileAttributes FileNameStr, attr
    GetTextAttributes
10:
ErrorTrap "applying attributes"
End Sub

Private Sub chkArchive_Click()
ApplyAttributes
End Sub

Private Sub chkHidden_Click()
ApplyAttributes
End Sub

Private Sub chkReadOnly_Click()
ApplyAttributes
End Sub

Private Sub chkSystem_Click()
ApplyAttributes
End Sub

Private Sub chkTemporary_Click()
ApplyAttributes
End Sub

Private Sub cmdLoad_Click()
    Dim strFile As String
    strFile = ShowCommonDlg(True, "", Me, "All Files (*)" & Chr(0) & "*" & Chr(0), "Load", 4096)
    If strFile <> "" Then
        FileNameStr = strFile
        txtFileName.Tag = 1
    Else
        If FileNameStr = "" Then
            txtFileName.Tag = 0
        Else
            txtFileName.Tag = 1
        End If
        Exit Sub
    End If
    GetFileInfo
    ToggleEnabled (True)
    Exit Sub
10:
    ErrorTrap "loading a file", FileNameStr
    txtFileName.Tag = 0
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo 10
    Dim iMsgBoxReturn As Integer
    If FileNameStr = "" Then
        CustomBox "There is no file loaded.", "Cannot set attributes to file because there is no file loaded.", vbExclamation, "", "", "OK"
        Exit Sub
    End If
    If GetAttr(FileNameStr) And vbReadOnly Then
        If CustomBox("Are you sure you want to delete this read-only file?", "If you choose Delete, this read-only file will be permanently deleted.", vbExclamation, "", "&Cancel", "&Delete") = 1 Then
            chkReadOnly.Value = Unchecked
            ApplyAttributes
        Else
            Exit Sub
        End If
    Else
        If CustomBox("Are you sure you want to delete this file?", _
        "If you choose Delete, this file will be permanently deleted.", _
        vbExclamation, "", "&Cancel", "&Delete") <> 1 Then Exit Sub
    End If
    Kill FileNameStr
            FileNameStr = ""
            txtFileName.Text = "No File"
            txtFileName.Tag = 0
            Uncheck
            lblModified.Caption = "--/--/---- --:--:-- --"
            lblSize.Caption = "0 Bytes"
            lblAttr.Caption = "No File Loaded"
            lblLocation.Caption = "\"
            lblLocation.Tag = "\"
            lblKind1.Caption = "Nonexistent file"
            lblFormat.Caption = ""
            ToggleEnabled (False)
10:
    ErrorTrap "deleting a file", FileNameStr
End Sub
Private Sub Uncheck()
On Error GoTo 10
    chkReadOnly.Value = 0
    chkArchive.Value = 0
    chkSystem.Value = 0
    chkHidden.Value = 0
    chkTemporary.Value = 0
10:
    ErrorTrap ""
End Sub
Private Sub cmdCopy_Click()
    On Error GoTo 10
    Dim ToFile As String
    If FileNameStr = "" Then
        CustomBox "There is no file loaded.", "Could not copy file because there is no file loaded.", vbExclamation, "", "", "OK"
        Exit Sub
    End If
    ToFile = ShowCommonDlg(False, "", Me, "All Files (*)" & Chr(0) & "*" & Chr(0), "Copy To...", 2)
    If ToFile = "" Then Exit Sub
    If Len(ToFile) = 0 Then Exit Sub
    FileCopy FileNameStr, ToFile
    Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
    ErrorTrap "copying a file", FileNameStr
End Sub
Private Sub cmdMove_Click()
   On Error GoTo 10
   Dim ToFile As String
    If FileNameStr = "" Then
        CustomBox "There is no file loaded.", "Could not move file because there is no file loaded.", vbExclamation, "", "", "OK"
        Exit Sub
    End If
    If GetAttr(FileNameStr) And vbReadOnly Then
        If CustomBox("Are you sure you want to move this read-only file?", "The file you tried to move is read-only. If you choose Move, it will no longer be read-only.", vbExclamation, "", "&Cancel", "&Move") = 1 Then
            chkReadOnly.Value = Unchecked
            ApplyAttributes
        Else
            Exit Sub
        End If
    End If
    ToFile = ShowCommonDlg(False, "", Me, "All Files (*)" & Chr(0) & "*" & Chr(0), "Move To...", 2)
    If ToFile = "" Then Exit Sub
    FileCopy FileNameStr, ToFile
    Kill FileNameStr
   FileNameStr = ToFile
   GetFileInfo
   Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
    ErrorTrap "moving a file", FileNameStr
End Sub


Private Sub cmdOK_Click()
On Error GoTo 10
If FileNameStr <> "" And txtFileName.Text <> "" Then txtFileName_KeyPress (13)
10:
ErrorTrap , FileNameStr
Unload Me
End Sub

Private Sub Command1_Click()
    If chkReadOnly.Value = Checked Then chkReadOnly.Value = Unchecked
    If chkArchive.Value = Checked Then chkArchive.Value = Unchecked
    If chkSystem.Value = Checked Then chkSystem.Value = Unchecked
    If chkHidden.Value = Checked Then chkHidden.Value = Unchecked
    If chkTemporary.Value = Checked Then chkTemporary.Value = Unchecked
    SetAttr FileNameStr, vbNormal
End Sub

Private Sub Form_Load()
On Error GoTo 10
Dim nofile As Boolean
If fMainForm.ActiveForm Is Nothing Then
nofile = True
txtFileName.Tag = 0
ToggleEnabled (False)
Exit Sub
End If
If fMainForm.ActiveForm.rtfText.FileName <> "" Then
    FileNameStr = fMainForm.ActiveForm.rtfText.FileName
    GetFileInfo
    txtFileName.Tag = 1
    lblSize.Caption = GetFileSize(FileNameStr) & " (" & FileLen(FileNameStr) & " bytes)"
Else
    nofile = True
    ToggleEnabled (False)
    lblSize.Caption = "[Current Document] " & LenB(fMainForm.ActiveForm.rtfText.TextRTF) & " bytes"
End If
10:
ErrorTrap "loading file management window"
End Sub
Private Sub GetFileInfo()
On Error GoTo 10
Dim FileNameExt As String
Dim FileExtInfo As String
Dim FileMIME As String
If InStr(1, FileNameStr, vbNullChar) <> 0 Then FileNameStr = Left(FileNameStr, InStr(1, FileNameStr, vbNullChar) - 1)
If Len(ParseFileName(FileNameStr)) < 40 Then
    txtFileName.Text = ParseFileName(FileNameStr)
Else
    txtFileName.Text = Left(ParseFileName(FileNameStr), 40) & "..."
End If
GetFileAttributes
GetTextAttributes
lblModified.Caption = FileDateTime(FileNameStr)
FileNameExt = Right$(FileNameStr, Len(FileNameStr) - InStrRev(FileNameStr, "."))
If InStr(1, ParseFileName(FileNameStr), ".") <> 0 Then
    lblFormat.Caption = UCase(FileNameExt)
        Select Case LCase(FileNameExt)
            Case "jpg", "jpe", "jpeg", "jfif", "bmp", "dib", "rle", "png", "tga", "tpic", "pntg", "tif", "tiff", "gif", "pict", "pct"
                FileExtInfo = UCase(FileNameExt) & " image"
            Case "txt", "text"
                FileExtInfo = "Plain text file"
            Case "wmf", "emf"
                FileExtInfo = "Metafile"
            Case "rtf"
                FileExtInfo = "text/rtf; Microsoft Rich Text Format"
            Case "wri"
                FileExtInfo = "Microsoft Windows Write file"
            Case "doc"
                FileExtInfo = "application/msword; Microsoft Word Document"
            Case "pdf"
                FileExtInfo = "application/pdf; Adobe Portable Document Format"
            Case "xml"
                FileExtInfo = "application/xml; W3C Extensible Markup Language format"
            Case "svg", "svgz"
                FileExtInfo = "image/svg+xml; W3C Scalable Vector Graphics vector image"
            Case "php", "shtml", "shtm", "jsp", "css", "html", "xhtml", "asp", "aspx"
                FileExtInfo = UCase(FileNameExt) & " Web source code"
            Case "exe"
                FileExtInfo = "Binary application"
            Case "mov", "qt", "mqv"
                FileExtInfo = "video/quicktime, Apple QuickTime container format"
            Case "aiff", "aif", "wav", "mp3", "m4a", "m4p", "ra", "rmvb", "au", "snd", "mid", "midi", "rmi", "aac"
                FileExtInfo = UCase(FileNameExt) & " Audio file"
            Case "mpg", "mpeg", "wmv", "m4v", "avi", "dv"
                FileExtInfo = UCase(FileNameExt) & " Video file"
            Case "tar", "jar", "gz", "bz2", "zip", "z", "lzh", "rar", "rpm", "cab", "7z", "001", "cpio", "deb", "sfx", "sit", "sitx"
                FileExtInfo = UCase(FileNameExt) & " archive"
            Case "prf"
                FileExtInfo = "Hyperwrite preferences file"
            Case "csv"
                FileExtInfo = "Comma Separated Values"
            Case "dll", "so"
                FileExtInfo = "Library"
            Case Else
                FileExtInfo = UCase(FileNameExt) & " file"
        End Select
        lblKind1.Caption = FileExtInfo
Else
    lblFormat.Caption = ""
    lblKind1.Caption = "Generic File; No extension"
End If
If Len(lblFormat.Caption) > 4 Then lblFormat.Caption = Left$(lblFormat.Caption, 3) & ".."
lblSize.Caption = GetFileSize(FileNameStr) & " (" & FileLen(FileNameStr) & " bytes)"
lblLocation.Tag = Left$(FileNameStr, InStrRev(FileNameStr, "\") - 1)
lblLocation.Caption = TrimLongWords(lblLocation.Tag, 63)
Exit Sub
10:
txtFileName.Text = "No file"
txtFileName.Tag = 0
ErrorTrap "getting file info", FileNameStr
End Sub

Private Function GetFileAttributes()
On Error GoTo 10
Dim FileAttr As Long
If FileNameStr = "" Then Exit Function
FileAttr = GetAttr(FileNameStr)
    If FileAttr And vbReadOnly Then
        chkReadOnly.Value = Checked
    End If
    If FileAttr And vbArchive Then
        chkArchive.Value = Checked
    End If
    If FileAttr And vbSystem Then
        chkSystem.Value = Checked
    End If
    If FileAttr And vbHidden Then
        chkHidden.Value = Checked
    End If
    If FileAttr And vbNormal Then
        Uncheck
    End If
    If FileAttr And vbTemporary Then
        chkTemporary.Value = Checked
    End If
10:
ErrorTrap "getting file attributes", FileNameStr
End Function

Private Function GetTextAttributes()
On Error GoTo 10
If FileNameStr = "" Then Exit Function
lblAttr.Caption = ""
If GetAttr(FileNameStr) And vbAlias Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Alias"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Alias"
    End If
End If
If GetAttr(FileNameStr) And vbArchive Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Archive"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Archive"
    End If
End If
If GetAttr(FileNameStr) And vbDirectory Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Directory"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Directory"
    End If
End If
If GetAttr(FileNameStr) And vbHidden Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Hidden"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Hidden"
    End If
End If
If GetAttr(FileNameStr) And vbNormal Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Normal"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Normal"
    End If
End If
If GetAttr(FileNameStr) And vbReadOnly Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Read-Only"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Read-Only"
    End If
End If
If GetAttr(FileNameStr) And vbSystem Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "System"
    Else
        lblAttr.Caption = lblAttr.Caption & ", System"
    End If
End If
If GetAttr(FileNameStr) And vbVolume Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Volume"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Volume"
    End If
End If
If GetAttr(FileNameStr) And vbCompressed Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Compressed"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Compressed"
    End If
End If
If GetAttr(FileNameStr) And vbTemporary Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Temporary"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Temporary"
    End If
End If
If lblAttr.Caption = "" Then lblAttr.Caption = "Normal"
10:
ErrorTrap "while getting file attributes", FileNameStr
End Function
Private Function GetFileSize(FileName) As String
    On Error GoTo 10
    If FileName = "" Then
        If fMainForm.ActiveForm Is Nothing Then Exit Function
        GetFileSize = FileLen(fMainForm.ActiveForm.rtfText.TextRTF) & " Bytes"
        Exit Function
    End If
    Dim sTemp As String
    sTemp = FileLen(FileName)
    If sTemp >= "1024" Then
        sTemp = CCur(sTemp / 1024) & " KB"
    Else
        If sTemp >= "1048576" Then
            sTemp = CCur(sTemp / (1024 * 1024)) & " MB"
        Else
            sTemp = CCur(sTemp) & " Bytes"
        End If
    End If
    GetFileSize = sTemp
10:
    If Err.Number = 0 Then Exit Function
    GetFileSize = "0 Bytes"
    ErrorTrap "calculating file size", FileNameStr
End Function
Private Function ToggleEnabled(TrueFalse As Boolean)
On Error GoTo 10
cmdMove.Enabled = TrueFalse
cmdCopy.Enabled = TrueFalse
cmdDelete.Enabled = TrueFalse
chkReadOnly.Enabled = TrueFalse
chkHidden.Enabled = TrueFalse
chkArchive.Enabled = TrueFalse
chkSystem.Enabled = TrueFalse
chkTemporary.Enabled = TrueFalse
10:
ErrorTrap ""
End Function

Private Sub lblLocation_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblLocation.BackColor = &H0&
lblLocation.ForeColor = RGB(255, 255, 0)
End Sub

Private Sub lblLocation_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblLocation.BackColor = &HFFFFFF
lblLocation.ForeColor = RGB(0, 0, 255)
DoEvents
Shell "explorer.exe " & lblLocation.Tag, vbNormalFocus
End Sub

Private Sub txtFileName_GotFocus()
txtFileName.Locked = txtFileName.Tag = 0
If txtFileName.Locked = False Then
    txtFileName.BorderStyle = 1
    txtFileName.SelStart = 0
    If InStrRev(txtFileName, ".") <> 0 Then
        txtFileName.SelLength = InStrRev(txtFileName, ".") - 1
    Else
        txtFileName.SelLength = Len(txtFileName)
    End If
End If
txtFileName.Tag = 1
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFileName.Locked = True Or txtFileName.Tag = 0 Then Exit Sub
    Dim strFile As String
    Dim strQuote As String
    strQuote = """"
    txtFileName.Text = Replace$(txtFileName.Text, "\", "")
    txtFileName.Text = Replace$(txtFileName.Text, "/", "")
    txtFileName.Text = Replace$(txtFileName.Text, "?", "")
    txtFileName.Text = Replace$(txtFileName.Text, ":", "")
    txtFileName.Text = Replace$(txtFileName.Text, "*", "")
    txtFileName.Text = Replace$(txtFileName.Text, "<", "")
    txtFileName.Text = Replace$(txtFileName.Text, ">", "")
    txtFileName.Text = Replace$(txtFileName.Text, "|", "")
    txtFileName.Text = Replace$(txtFileName.Text, strQuote, "")
    txtFileName.BorderStyle = 0
    strFile = Left(FileNameStr, InStrRev(FileNameStr, "\")) & txtFileName.Text
    Name FileNameStr As strFile
    FileNameStr = strFile
    txtFileName.Tag = 1
    cmdLoad.SetFocus
End If
End Sub
