VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmDocument 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Document"
   ClientHeight    =   3660
   ClientLeft      =   2220
   ClientTop       =   2130
   ClientWidth     =   5595
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   5595
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdrag 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Drag"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3705
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   6535
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      MaxLength       =   2000000
      Appearance      =   0
      TextRTF         =   $"frmDocument.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
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
Public bChanged As Boolean

'AutoSave
Public CurrentStart As Long
Public lSaveStart As Long
Public bAutoSave As Boolean

'Tables/Elastic Tables
Dim bStepOne As Boolean
Dim XLng As Long, YLng As Long

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo 10
Dim iMsgBoxReturn As Integer
If bChanged <> False Then
Me.SetFocus
If rtfText.FileName <> "" Then
    If GetAttr(rtfText.FileName) And vbReadOnly Then
        iMsgBoxReturn = CustomBox("This file has unsaved changes, but is read-only. Do you want to save this file to another location?", "This file can only be saved to another location. If you don" & sApostrophe & "t save, your changes will be lost.", vbExclamation, "&Don" & sApostrophe & "t Save", "&Cancel", "&Save As...")
        If iMsgBoxReturn = 1 Then fMainForm.mnuFileSaveAs_Click
        If iMsgBoxReturn = 2 Then Cancel = 1
        Exit Sub
    End If
End If
    iMsgBoxReturn = CustomBox("Do you want to save changes to " & TrimLongWords(Replace(Me.Caption, "*", "")) & " before closing?", "If you don" & sApostrophe & "t save, your changes will be lost.", vbExclamation, "&Don" & sApostrophe & "t Save", "&Cancel", "&Save")
    If iMsgBoxReturn = 1 Then fMainForm.mnuFileSave_Click
    If iMsgBoxReturn = 2 Then
        Cancel = 1
        Exit Sub
    End If
End If
fMainForm.StatusBar.Panels(1).Text = "Ready"
10:
ErrorTrap "closing a child window", rtfText.FileName
End Sub

Private Sub rtfText_Change()
On Error GoTo 10
If bNoStatus = True Then Exit Sub
If bAutoSave = True Then
    CurrentStart = CurrentStart + 1
    If lSaveStart >= CurrentStart Then
        fMainForm.mnuFileSave_Click
        CurrentStart = 0
        Exit Sub
    End If
End If
    bChanged = True
    If fMainForm.mnuEditUndoReplace.Enabled = True Then
        fMainForm.mnuEditUndoReplace.Enabled = False
        rtfText.Tag = ""
    End If
    Me.Caption = ParseFileName(rtfText.FileName) & "*"
    fMainForm.mnuFileSave.Enabled = True
10:
End Sub

Private Sub rtfText_DblClick()
bRubberBand = False
fMainForm.mnuTableElastic.Checked = False
End Sub

Private Sub rtfText_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo 10
Dim strSymbol As String
Dim lngStart As Long
Dim lngLength As Long
If KeyDown(vbKeyControl) Then
    If Shift <> 2 Then
        Select Case KeyCode
            Case 49 To 53
                fMainForm.ChangeFont (KeyCode - 49)
            Case 188
                DoWords True, True
                KeyCode = 0
            Case 190
                DoWords False, True
                KeyCode = 0
            Case 83
                fMainForm.mnuFileSaveAs_Click
                KeyCode = 0
        End Select
    Else
        Select Case KeyCode
            
            Case 32
                rtfText.SelText = Chr(160) 'Non-breaking space
                KeyCode = 0
            Case 86, 88
                If bLiveWC = True Then
                    fMainForm.tmrLiveWC.Enabled = False
                    fMainForm.tmrLiveWC.Enabled = True
                End If
        End Select
    End If
End If
If bAltMode = True Then
Dim lngCode As Long
    If KeyCode = 190 And Shift = 0 Then lngCode = 12290
    If KeyDown(vbKeyControl) = True Then
        If KeyCode = 186 Then lngCode = 12300
        If KeyCode = 222 Then lngCode = 12301
    End If
    If lngCode <> 0 Then
        KeyCode = 0
        rtfText.SelText = "\u" & lngCode & "?"
        lngStart = rtfText.SelStart
        rtfText.TextRTF = Replace(rtfText.TextRTF, "\\u" & lngCode & "?", "\u" & lngCode & "?")
        rtfText.SelStart = lngStart - Len(CStr(lngCode))
    End If
End If
If bRealSymbols = True Then
    Static bJustChanged As Boolean
    Static btChar As Byte
    Static bSmileSymbol As Boolean 'SymbolMatic for Smileys: 0 = ":"
    Static bDotSymbol(1) As Boolean 'SymbolMatic for Ellipsis: 0 and 1 = "."
    Static bMarkSymbol(2) As Boolean 'SymbolMatic for Copyright, Trademark, and Registered TradeMark
    Static bRealSymbol(2) As Boolean 'SymbolMatic for Arrows: 0 = "<", 1 = ">", 2 = "-"
    Static lngSymbolPos As Long
    If bJustChanged = True Then
        If lngSymbolPos = rtfText.SelStart Then
            If KeyCode = 8 Then
                If Shift = 0 Then
                    KeyCode = 0
                    rtfText.SelStart = rtfText.SelStart - 1
                    rtfText.SelLength = 1
                    Select Case btChar
                        Case 0 'Copyright
                            rtfText.SelText = "(C)"
                        Case 1 'Registered Trademark
                            rtfText.SelText = "(R)"
                        Case 2 'Trademark
                            rtfText.SelText = "(TM)"
                        Case 3 'Smiley
                            rtfText.SelText = ":)"
                        Case 4 'Ellipsis
                            rtfText.SelText = "..."
                        Case 5 '<<
                            rtfText.SelText = "<<"
                        Case 6 '<-
                            rtfText.SelText = "<-"
                        Case 7 '->
                            rtfText.SelText = "->"
                        Case 8 '>>
                            rtfText.SelText = ">>"
                    End Select
                End If
            End If
        Else
            lngSymbolPos = 0
        End If
        bJustChanged = False
        btChar = 0
    End If
    If bMarkSymbol(1) = True Then
        If Shift = 1 Then
            If KeyCode = 48 Then
                    strSymbol = Mid$(rtfText.Text, rtfText.SelStart, 1)
                    If strSymbol = "C" Then
                        KeyCode = 0
                        rtfText.SelStart = rtfText.SelStart - 2
                        rtfText.SelLength = 2
                        rtfText.SelText = Chr(169)
                        bJustChanged = True
                        lngSymbolPos = rtfText.SelStart
                        btChar = 0
                    End If
                    If strSymbol = "R" Then
                        KeyCode = 0
                        rtfText.SelStart = rtfText.SelStart - 2
                        rtfText.SelLength = 2
                        rtfText.SelText = Chr(174)
                        bJustChanged = True
                        lngSymbolPos = rtfText.SelStart
                        btChar = 1
                    End If
                    bMarkSymbol(2) = False
                    bMarkSymbol(1) = False
                    bMarkSymbol(0) = False
            End If
            If KeyCode = 77 Then
                bMarkSymbol(2) = True
            Else
                bMarkSymbol(2) = False
                bMarkSymbol(1) = False
                bMarkSymbol(0) = False
            End If
        End If
    End If
    If bMarkSymbol(2) = True Then
        If Shift = 1 Then
            If KeyCode = 48 Then
                strSymbol = Mid$(rtfText.Text, rtfText.SelStart - 1, 2)
                If strSymbol = "TM" Then
                    KeyCode = 0
                    rtfText.SelStart = rtfText.SelStart - 3
                    rtfText.SelLength = 3
                    rtfText.SelText = Chr(153)
                    bJustChanged = True
                    lngSymbolPos = rtfText.SelStart
                    btChar = 2
                End If
                bMarkSymbol(2) = False
                bMarkSymbol(1) = False
                bMarkSymbol(0) = False
            End If
        End If
    End If
    If bSmileSymbol = True Then
        If KeyCode = 48 Then
            If Shift = 1 Then
                KeyCode = 0
                rtfText.SelStart = rtfText.SelStart - 1
                lngStart = rtfText.SelStart
                rtfText.SelLength = 1
                rtfText.SelText = "\" & rtfText.SelText
                rtfText.SelStart = lngStart
                rtfText.SelLength = 2
                rtfText.SelRTF = Replace(rtfText.SelRTF, "\" & rtfText.SelText, "\u9786?")
                bSmileSymbol = False
                bJustChanged = True
                lngSymbolPos = rtfText.SelStart
                btChar = 3
            End If
        Else
            bSmileSymbol = False
        End If
    End If
    If KeyCode = 186 Then
        If Shift = 1 Then bSmileSymbol = True
    End If
    If bMarkSymbol(0) = True Then
        If Shift = 1 Then
            If KeyCode = 67 Or KeyCode = 82 Or KeyCode = 84 Then
                bMarkSymbol(1) = True
            Else
                bMarkSymbol(1) = False
                bMarkSymbol(0) = False
            End If
        End If
    End If
    If KeyCode = 57 Then
        If Shift = 1 Then bMarkSymbol(0) = True
    End If
    If bDotSymbol(1) = True Then 'Have two periods been entered already?
        If Shift = 0 Then 'Make sure a period was entered instead of a right bracket
            If KeyCode = 190 Then 'Period
                KeyCode = 0
                rtfText.SelStart = rtfText.SelStart - 2
                rtfText.SelLength = 2
                rtfText.SelText = Chr(133) 'Ellipsis
                bJustChanged = True
                lngSymbolPos = rtfText.SelStart
                btChar = 4
            End If
        End If
        bDotSymbol(1) = False 'Reset period traces
        bDotSymbol(0) = False
    End If
    If bDotSymbol(0) = True Then 'Has one period been entered already?
        If Shift = 0 Then
            If KeyCode = 190 Then bDotSymbol(1) = True 'If the current keystroke is a period, two periods have been entered.
        Else
            bDotSymbol(0) = False
        End If
    End If
    If bRealSymbol(0) = True Then
      If Shift = 1 Then
        If KeyCode = 188 Then
            KeyCode = 0
            rtfText.SelStart = rtfText.SelStart - 1
            rtfText.SelLength = 1
            rtfText.SelText = "«"
            bRealSymbol(0) = False
            bJustChanged = True
            lngSymbolPos = rtfText.SelStart
            btChar = 5
        End If
      End If
        If KeyCode = 189 Then
            If bRealSymbol(0) = True Then
                KeyCode = 0
                rtfText.SelStart = rtfText.SelStart - 1
                rtfText.SelLength = 1
                rtfText.SelText = Chr(27) '<-
                bRealSymbol(0) = False
                bJustChanged = True
                lngSymbolPos = rtfText.SelStart
                btChar = 6
            End If
        End If
    End If
    If bRealSymbol(2) = True Then
            bRealSymbol(0) = False
            bRealSymbol(1) = False
              If Shift = 1 Then
                If KeyCode = 190 Then
                    KeyCode = 0
                    rtfText.SelStart = rtfText.SelStart - 1
                    rtfText.SelLength = 1
                    rtfText.SelText = Chr(26) '->
                    bRealSymbol(2) = False
                    bJustChanged = True
                    lngSymbolPos = rtfText.SelStart
                    btChar = 7
                End If
              End If
        End If
    If bRealSymbol(1) = True Then
        If KeyCode = 190 Then
          If Shift = 1 Then
            KeyCode = 0
            rtfText.SelStart = rtfText.SelStart - 1
            rtfText.SelLength = 1
            rtfText.SelText = "»"
            bRealSymbol(1) = False
            bJustChanged = True
            lngSymbolPos = rtfText.SelStart
            btChar = 8
          End If
        End If
    End If
     If Shift = 1 Then
        bRealSymbol(0) = KeyCode = 188
        bRealSymbol(1) = KeyCode = 190
     Else
        bRealSymbol(2) = KeyCode = 189 Or KeyCode = 16
        bDotSymbol(0) = KeyCode = 190
     End If
    'If bLiveWC = False Then fMainForm.StatusBar.Panels(1).Text = "ab| (" & KeyCode & ")"
    If KeyDown(vbKeyControl) = True Then Exit Sub
      If KeyCode = 222 Then
        If rtfText.SelStart = 0 Then
            If Shift = 0 Then
                rtfText.SelText = Chr(145) 'Left Single Quote
            Else
                rtfText.SelText = Chr(147) 'Left Double Quote
            End If
            KeyCode = 0
        Exit Sub
      End If
            If InStr("| |" & vbTab & "|" & vbLf & "|" & vbCr, Mid$(rtfText.Text, rtfText.SelStart, 1)) <> 0 Then
                If Shift = 1 Then
                    rtfText.SelText = Chr(147)
                Else
                    rtfText.SelText = Chr(145) 'Left Single Quote
                End If
                KeyCode = 0
            Else
                If Shift = 1 Then
                    rtfText.SelText = Chr(148)
                Else
                    rtfText.SelText = sApostrophe 'Right Single Quote
                End If
                KeyCode = 0
        End If ' SelText
       End If 'KeyCode
End If 'SymbolMatic
'Exit Sub
10:
ErrorTrap "KeyDown"
End Sub

Private Sub rtfText_KeyPress(KeyAscii As Integer)
If bLiveWC = True Then
    fMainForm.tmrLiveWC.Enabled = False
    fMainForm.tmrLiveWC.Enabled = True
End If
End Sub

Private Sub rtfText_KeyUp(KeyCode As Integer, Shift As Integer)
If bLiveWC = False Then fMainForm.StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu fMainForm.mnuRightClick
End If
If bRubberBand = True Then
bStepOne = True
XLng = x
txtdrag.Visible = True
txtdrag.Top = Y
txtdrag.Left = x
End If
End Sub

Private Sub rtfText_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bLiveWC = False Then fMainForm.StatusBar.Panels(1).Text = "Ready"
ScaleMode = vbTwips
If bRubberBand = True Then
    If bStepOne = True Then
    If txtdrag.Visible = False Then Exit Sub
        'txtdrag.Left = x + 100
        txtdrag.Height = 100
        If x < XLng Then Exit Sub
        txtdrag.Width = x - XLng
        ScaleMode = vbInches
        txtdrag.Text = txtdrag.Width & " inches"
        ScaleMode = vbTwips
    End If
End If
End Sub

Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bRubberBand = True Then
ScaleMode = vbTwips
rtfText.MousePointer = 3
txtdrag.Visible = False
Dim TableWidth As Integer
TableWidth = x - XLng
If TableWidth = 0 Then Exit Sub
lCellWidth(0) = TableWidth
CreateTable 1, 1
bStepOne = False
End If
End Sub

Private Sub rtfText_SelChange()
On Error GoTo 10
    If bNoStatus = True Or InStr(rtfText.SelRTF, "{\pict") <> 0 Or InStr(rtfText.SelRTF, "{\object") <> 0 Then Exit Sub
    fMainForm.tbFormat.Buttons(1).Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(2).Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(3).Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(4).Value = IIf(rtfText.SelStrikeThru, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(6).Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(7).Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(8).Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(12).Value = IIf(InStr(1, rtfText.SelRTF, "\super") <> 0, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(13).Value = IIf(InStr(1, rtfText.SelRTF, "\sub") <> 0, tbrPressed, tbrUnpressed)
    fMainForm.tbFormat.Buttons(10).Value = IIf(InStr(1, rtfText.SelRTF, "\pntext") <> 0, tbrPressed, tbrUnpressed)
    If rtfText.SelColor <> vbNull Then
    fMainForm.shpSwatch.BackColor = rtfText.SelColor
    End If
    fMainForm.StatusBar.Panels(2).Text = "Ln " & rtfText.GetLineFromChar(rtfText.SelStart) + 1 & "  Pos " & rtfText.SelStart & "  Sel " & rtfText.SelLength & " "
    If rtfText.SelFontName <> vbNull Then
        fMainForm.cboFontFace.Text = rtfText.SelFontName
        fMainForm.txtPreview.Text = " " & fMainForm.cboFontFace.Text
    Else
        fMainForm.cboFontFace.Text = ""
        fMainForm.txtPreview.Text = ""
    End If
    If rtfText.SelFontSize <> vbNull Then
        fMainForm.cboFontSize.Text = rtfText.SelFontSize
        fMainForm.tbFormat.Buttons(14).Enabled = True
        fMainForm.tbFormat.Buttons(15).Enabled = True
    Else
        fMainForm.cboFontSize.Text = ""
        fMainForm.tbFormat.Buttons(14).Enabled = False
        fMainForm.tbFormat.Buttons(15).Enabled = False
    End If
    If bNormal = False Then lRightMargin = rtfText.RightMargin
10:
End Sub

Private Sub Form_Load()
    bStepOne = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 180, 180, Me.ScaleWidth - 200, Me.ScaleHeight - 200
End Sub
