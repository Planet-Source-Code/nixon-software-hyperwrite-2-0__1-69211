Attribute VB_Name = "modMain"
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
        ' Hyperwrite from NIXON                                  '
        ' Copyright (C) 2004-2007 NIXON Software Corporation.    '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
        ' You may use this code freely in your own applications. '
        ' If you are distributing your code/application(s), it   '
        ' would be greatly appreciated if you credit NIXON in    '
        ' your About dialog. Please note that portions of this   '
        ' code may belong to other parties. For more details,    '
        ' please view the About dialog.                          '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As KeyCodeConstants) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Public fMainForm As frmMain
Public Const sApostrophe = "’"
Public Const WM_CLEAR = &H303
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302
Public Const EM_GETLINECOUNT = &HBA
Public bAltMode As Boolean
Public lCellWidth(49) As Long
Public bLiveWC As Boolean
Public bNormal As Boolean
Public bRealSymbols As Boolean
Public bNoStatus As Boolean
Public bRubberBand As Boolean
Public intMsgReturn As Integer
Public btDocumentCount As Byte
Public lRightMargin As Long
Private strPref As String

'CommonDialog without OCX
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Public sFileName As String

Sub Main()
On Error Resume Next
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    LoadPrefFile
Load fMainForm
fMainForm.Show
Unload frmSplash
bRealSymbols = DoPrefs(0, "DefSymbolMatic") = 1
End Sub
Public Function ChangeRTF(sAttr As String)
Dim lngLength As Long
Dim lngSel As Long
lngLength = fMainForm.ActiveForm.rtfText.SelLength
lngSel = fMainForm.ActiveForm.rtfText.SelStart
bNoStatus = True
fMainForm.ActiveForm.rtfText.SelText = "\" & fMainForm.ActiveForm.rtfText.SelText
fMainForm.ActiveForm.rtfText.TextRTF = Replace(fMainForm.ActiveForm.rtfText.TextRTF, "\\" & fMainForm.ActiveForm.rtfText.SelText, "\" & sAttr & fMainForm.ActiveForm.rtfText.SelText)
bNoStatus = False
fMainForm.ActiveForm.rtfText.SelStart = lngSel
fMainForm.ActiveForm.rtfText.SelLength = lngLength
End Function
Public Function DoWords(Optional bReverse As Boolean = False, Optional bSelect As Boolean = False) As Boolean
On Error GoTo 10
Dim lngSp(1) As Long, lngDiff As Long
If bReverse = False Then
    lngSp(0) = InStr(fMainForm.ActiveForm.rtfText.SelStart + 1, fMainForm.ActiveForm.rtfText.Text, " ", vbTextCompare) + 1
    lngSp(1) = InStr(lngSp(0) + 1, fMainForm.ActiveForm.rtfText.Text, " ", vbTextCompare)
    If lngSp(1) = 0 Then
        DoWords = False
        fMainForm.ActiveForm.rtfText.SelStart = lngSp(0) - 1
        fMainForm.ActiveForm.rtfText.SelLength = Len(fMainForm.ActiveForm.rtfText.Text) - lngSp(0) + 1
        Exit Function
    End If
    lngDiff = lngSp(1) - lngSp(0)
    fMainForm.ActiveForm.rtfText.SelStart = lngSp(0) - 1
    If bSelect = True Then fMainForm.ActiveForm.rtfText.SelLength = lngDiff
Else
    If fMainForm.ActiveForm.rtfText.SelStart = 0 Then
        lngSp(0) = InStrRev(fMainForm.ActiveForm.rtfText.Text, " ")
        fMainForm.ActiveForm.rtfText.SelStart = lngSp(0)
        fMainForm.ActiveForm.rtfText.SelLength = Len(fMainForm.ActiveForm.rtfText.Text) - lngSp(0)
        Exit Function
    End If
    lngSp(0) = InStrRev(fMainForm.ActiveForm.rtfText.Text, " ", fMainForm.ActiveForm.rtfText.SelStart, vbBinaryCompare)
    lngSp(1) = InStrRev(fMainForm.ActiveForm.rtfText.Text, " ", lngSp(0) - 1, vbBinaryCompare)
    lngDiff = lngSp(0) - lngSp(1)
    fMainForm.ActiveForm.rtfText.SelStart = lngSp(1)
    If bSelect = True Then fMainForm.ActiveForm.rtfText.SelLength = lngDiff - 1
End If
    DoWords = True
10:
End Function
Public Function ShowCommonDlg(bShowOpen As Boolean, strDefExt As String, hwndOwner As Form, strFilter As String, _
                            Optional strTitle As String = "", Optional lngFlags As Long = 0) As String
    On Error Resume Next
    Dim OpenFile As OPENFILENAME, lReturn As Long
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hwndOwner.hwnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = strFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrDefExt = strDefExt
    'OpenFile.lpstrInitialDir = "%HOMEPATH%\My Documents"
    OpenFile.lpstrTitle = strTitle
    OpenFile.flags = lngFlags
    If bShowOpen = True Then
        lReturn = GetOpenFileName(OpenFile)
    Else
        lReturn = GetSaveFileName(OpenFile)
    End If
    If lReturn <> 0 Then ShowCommonDlg = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
End Function
Public Function KeyDown(ByVal vKey As KeyCodeConstants) _
    As Boolean
   KeyDown = GetAsyncKeyState(vKey) And &H8000
End Function
Public Function GetLineCount(Control_hWnd As Long) As Long
On Error GoTo 10
    GetLineCount = SendMessage(Control_hWnd, EM_GETLINECOUNT, True, 0&)
10:
End Function
Public Function GetLastFontNum() As Integer
If fMainForm.ActiveForm Is Nothing Then Exit Function
    Dim lngFontTablePos As Long, lngEndofTablePos As Long
    Dim lngLastFontPos As Long, lngLastFontSlash As Long
    With fMainForm.ActiveForm.rtfText
        lngFontTablePos = InStr(1, .TextRTF, "{\fonttbl")
        lngEndofTablePos = InStr(lngFontTablePos, .TextRTF, "}}")
        lngLastFontPos = InStrRev(.TextRTF, "{\f", lngEndofTablePos)
        lngLastFontSlash = InStr(lngLastFontPos + 3, .TextRTF, "\")
        GetLastFontNum = Mid(.TextRTF, lngLastFontPos + 3, lngLastFontSlash - lngLastFontPos - 3)
    End With
End Function
Public Function ParseFontTable(intFont As Integer) As String
If fMainForm.ActiveForm Is Nothing Then Exit Function
    Dim lngFontPos As Long, lngFontEnd As Long, lngSpacePos As Long
    With fMainForm.ActiveForm.rtfText
        lngFontPos = InStr(1, .TextRTF, "{\f" & intFont)
        lngSpacePos = InStr(lngFontPos, .TextRTF, " ")
        lngFontEnd = InStr(lngSpacePos, .TextRTF, ";}")
        ParseFontTable = Mid(.TextRTF, lngSpacePos + 1, lngFontEnd - lngSpacePos - 1)
    End With
End Function
Public Function WordCount(strText As String) As Long
If fMainForm.ActiveForm Is Nothing Then CustomBox "No windows open", "Could not complete your request because there are no windows open.", vbExclamation, "", "", "&OK": Exit Function
    On Error Resume Next
    fMainForm.ActiveForm.MousePointer = vbHourglass
    DoEvents
    If strText = "" Then
        WordCount = 0
        Exit Function
    End If
    Dim counter As Long
    Dim words As Long
    Dim sum As Long
    counter = 0
    words = 0
    sum = 0
    strText = Trim(strText)
    fMainForm.StatusBar.Style = sbrSimple
    fMainForm.StatusBar.SimpleText = "Counting words..."
    strText = Replace(strText, vbNewLine, " ")  'Replace all tabs, new lines, and non-breaking spaces
    strText = Replace(strText, vbTab, " ")      'with a space so that you can count them later
    strText = Replace(strText, Chr(160), " ")
    Do While InStr(1, strText, "  ") <> 0
        strText = Replace(strText, "  ", " ")
    Loop
    strText = Trim(strText)
    fMainForm.StatusBar.Style = sbrNormal
    fMainForm.StatusBar.SimpleText = "Ready"
    WordCount = FindOccurrences(strText, " ") + 2
    fMainForm.ActiveForm.MousePointer = vbIbeam
End Function
Public Function TrimLongWords(strString As String, Optional intLen As Integer = 40) As String
    If Len(strString) > intLen Then
        TrimLongWords = Left(strString, intLen) & "..."
    Else
        TrimLongWords = strString
    End If
End Function
Public Sub ErrorTrap(Optional strInfo As String = "", Optional strFileName As String = "[unspecified]")
If Err.Number = 0 Then Exit Sub
fMainForm.StatusBar.Style = sbrSimple
fMainForm.StatusBar.SimpleText = "Error (" & Err.Number & ")"
    Select Case Err.Number
    Case 53
        CustomBox "The file " & ParseFileName(strFileName) & " could not be found.", "Please check on the spelling. The file may have been moved, renamed, or deleted.", vbCritical, "", "", "OK"
    Case 61
    CustomBox "The document " & Chr(147) & ParseFileName(strFileName) & Chr(148) _
        & " could not be saved because the " & Chr(147) & _
        GetDrive(strFileName, True) & Chr(148) & " is full.", "Try deleting documents from " & _
        Chr(147) & GetDrive(strFileName, False) & Chr(148) & " or saving the document on another disc." _
        , vbCritical, "", "", "OK"
    Case 71
    CustomBox "The document " & Chr(147) & ParseFileName(strFileName) & Chr(148) _
        & " could not be  because the " & Chr(147) & _
        GetDrive(strFileName, True) & Chr(148) & " is not ready.", "Check if the drive is open and the disc is inside the drive.", vbCritical, "", "", "OK"
    Case 72
    CustomBox "The file " & Chr(147) & ParseFileName(strFileName) & Chr(148) _
        & " could not be accessed because of a file I/O error.", "The file you were trying to access is corrupt. Please contact your system administrator for assistance.", vbCritical, "", "", "OK"
    Case 75
    CustomBox "The file " & Chr(147) & ParseFileName(strFileName) & Chr(148) _
        & " on " & GetDrive(strFileName, True) & " could not be accessed or is invalid.", "This can happen if the file is read-only or in an invalid format.", vbCritical, "", "", "OK"
    Case 57
    CustomBox "The " & Chr(147) & GetDrive(strFileName, True) & Chr(148) _
        & " could not be accessed because of a device I/O error.", "There is a problem with the device you were trying to save to. Please contact your system administrator for assistance.", vbCritical, "", "", "OK"
    Case Else
    If strInfo = "" Then
        CustomBox "Error " & Err.Number & " has occured. (Unknown location)", "Description: " & _
        Err.Description, vbCritical, "", "", "OK"
    Else
        CustomBox "Error " & Err.Number & " occured while " & strInfo & ".", "Description: " & _
        Err.Description, vbCritical, "", "", "OK"
    End If
    End Select
fMainForm.StatusBar.Style = sbrNormal
fMainForm.StatusBar.Panels(1).Text = "Ready"
End Sub
Private Function GetDrive(sFile As String, bVerbose As Boolean) As String
    If Left(sFile, 2) <> "\\" Then
        If bVerbose = True Then
            GetDrive = "drive " & Left(sFile, 2)
        Else
            GetDrive = Left(sFile, 2)
        End If
    Else
        If bVerbose = True Then
            GetDrive = "computer " & Chr(147) & Mid(sFile, 3, InStr(3, sFile, "\") - 3) & Chr(148)
        Else
            GetDrive = Mid(sFile, 3, InStr(3, sFile, "\") - 3)
        End If
    End If
End Function
Public Function SelectLine(rtfTheRtf As RichTextBox, lngLine As Long) As Boolean
If fMainForm.ActiveForm Is Nothing Then Exit Function
  On Error Resume Next
  Dim lngPos As Integer, lngLineCount As Long, blnFound As Boolean
  Dim blnEnd As Boolean, lngLineStart As Long
  'Start at beginning of text
  rtfTheRtf.SelStart = 0
   
  lngLineCount = 0
  blnFound = False
  lngPos = 0
  blnEnd = False
  'Go through the text until we find the right line or we hit the end of the
  'text or there are no more lines
  While lngLineCount < lngLine And rtfTheRtf.SelStart < Len(rtfTheRtf.Text) And _
        Not blnEnd
    'Save current position
    rtfTheRtf.SelStart = lngPos
    'Save position of first char of the current line
    lngLineStart = lngPos
    'Select text until end of line
    rtfTheRtf.Span vbCrLf, True, True
    'Span() does not advance the position so we have to do it manually
    rtfTheRtf.UpTo vbCrLf, True, False
    'If position has not moved, there aren't anymore lines
    If rtfTheRtf.SelStart = lngPos Then
      blnEnd = True
    Else
      'Count lines
      lngLineCount = lngLineCount + 1
      'Check if we found the right one
      If lngLineCount = lngLine Then blnFound = True
    End If
     
    'Advance position to the next line (over CRLF)
    lngPos = rtfTheRtf.SelStart + 2
  Wend
   
  'When the line is found then select it
  '(we have to do it again because UpTo() clears the selection)
  If blnFound Then
    'Select the line
    rtfTheRtf.SelStart = lngLineStart
    rtfTheRtf.Span vbCrLf, True, True
  End If
   
  SelectLine = blnFound
End Function
Public Function CreateTable(Cells As Integer, Rows As Integer)
Dim iString(5) As String
Dim tempString(3) As String
Dim i As Integer
fMainForm.ActiveForm.ScaleMode = vbPixels
iString(0) = "\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw10 "
iString(1) = "\trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 "
iString(2) = "\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx" & CLng(lCellWidth(0))
    If Cells > 1 Then
        For i = 1 To Cells - 1
            If i <> 1 Then
                iString(3) = iString(3) & "\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx" & CLng(lCellWidth(i)) + CLng(lCellWidth(i - 1))
            Else
                iString(3) = iString(3) & "\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx" & CLng(lCellWidth(i))
            End If
        tempString(0) = tempString(0) & "\cell"
        Next
    End If
For i = 0 To Rows - 1
    tempString(1) = tempString(1) & "\row"
Next
iString(4) = "\pard\intbl\f0\fs24" & "\cell" & tempString(0) & tempString(1) & vbNewLine & "\pard }"
For i = 0 To 5
    iString(5) = iString(5) + vbNewLine + iString(i)
Next
fMainForm.ActiveForm.rtfText.SelText = "\tr" & fMainForm.ActiveForm.rtfText.SelText
fMainForm.ActiveForm.rtfText.TextRTF = Replace(fMainForm.ActiveForm.rtfText.TextRTF, "\\tr" & fMainForm.ActiveForm.rtfText.SelText, iString(5))
End Function
Public Function CustomBox(sPrompt As String, sInfo As String, bStyle As VbMsgBoxStyle, _
        Optional sBt1 As String, Optional sBt2 As String, Optional sBt3 As String, Optional _
        btDefButton As Byte) As Integer
Dim i As Integer
With frmDialog
    .cmdButton(2).Caption = sBt1
    .cmdButton(1).Caption = sBt2
    .cmdButton(0).Caption = sBt3
    For i = 0 To 2
        If .cmdButton(i).Caption = "" Then
            .cmdButton(i).Visible = False
        Else
            .cmdButton(i).Visible = True
        End If
    Next
    Select Case bStyle
    Case vbCritical
        i = 3
    Case vbExclamation
        i = 2
    Case vbInformation
        i = 1
    Case Else
        i = 1
    End Select
    .lblMsg.Caption = sPrompt
    .lblInfo.Caption = sInfo
    .imgIcon.Picture = .imgList.ListImages(i).Picture
    .imgIcon.Left = 24 * Screen.TwipsPerPixelX
    .imgIcon.Top = 15 * Screen.TwipsPerPixelY
    .lblMsg.Top = 15 * Screen.TwipsPerPixelY
    .lblMsg.Left = .imgIcon.Left + .imgIcon.Width + (16 * Screen.TwipsPerPixelX)
    .lblMsg.Width = .ScaleWidth - .lblMsg.Left - (24 * Screen.TwipsPerPixelX)
    .lblInfo.Left = .lblMsg.Left
    .lblInfo.Width = .lblMsg.Width
    .lblInfo.Top = .lblMsg.Top + .lblMsg.Height + (8 * Screen.TwipsPerPixelY)
    .cmdButton(0).Left = (.ScaleWidth - 24 * Screen.TwipsPerPixelX) - .cmdButton(0).Width
    .cmdButton(1).Left = .cmdButton(0).Left - (12 * Screen.TwipsPerPixelX) - .cmdButton(1).Width
    .cmdButton(2).Left = .lblInfo.Left
    .cmdButton(0).Top = .lblInfo.Top + .lblInfo.Height + .cmdButton(0).Height
    .cmdButton(1).Top = .cmdButton(0).Top
    .cmdButton(2).Top = .cmdButton(0).Top
    .Height = .cmdButton(0).Top + .cmdButton(0).Height + (20 * Screen.TwipsPerPixelY) + (.Height - .ScaleHeight)
    Select Case btDefButton
        Case 1
            .cmdButton(0).Default = True
        Case 2
            .cmdButton(1).Default = True
        Case 3
            .cmdButton(2).Default = True
        Case Else
            .cmdButton(0).Default = True
    End Select
    .Show vbModal
End With
CustomBox = intMsgReturn
End Function
Public Function FindOccurrences(strFind As String, strMatch As String) As Long
Dim lngPos As Long
Dim lngCount As Long
lngPos = 1
lngCount = -1
Do While lngPos <> 0
    lngPos = InStr(lngPos + 1, strFind, strMatch)
If lngPos <> 0 Then lngCount = lngCount + 1
Loop
FindOccurrences = lngCount
End Function
Public Function ParseFileName(sFileIn As String) As String
If sFileIn = "" Then
    If btDocumentCount = 1 Then
        ParseFileName = "untitled"
    Else
        ParseFileName = "untitled " & btDocumentCount
    End If
    Exit Function
End If
Dim i As Integer
If InStrRev(sFileIn, "\") = 0 Then
    ParseFileName = sFileIn
Else
    ParseFileName = Right$(sFileIn, Len(sFileIn) - InStrRev(sFileIn, "\"))
End If
End Function

Private Function LoadPrefFile()
On Error GoTo 10
Dim FileNum%
FileNum% = FreeFile()
Open App.Path & "\prefs.prf" For Input As #FileNum%
strPref = Input(LOF(FileNum%), #FileNum%)
Close #FileNum%
10:
If Err.Number = 53 Then
    ResetPrefs
    SavePrefFile
End If
End Function

Public Sub ResetPrefs()
On Error Resume Next
strPref = "SaveWorkspace:<0>" & vbNewLine & _
                "DefSymbolMatic:<1>" & vbNewLine & _
                "AutoReplaceStraightQuotes:<0>" & vbNewLine & _
                "ImportPictures:<1>" & vbNewLine & _
                "RecentFiles:<1>" & vbNewLine & _
                "WarnTextFormat:<1>" & vbNewLine & _
                "StatusBarFind:<1>" & vbNewLine & _
                "ShowCoolbar:<1>" & vbNewLine & _
                "ShowToolbar:<1>" & vbNewLine & _
                "ShowFormatBar:<1>" & vbNewLine & _
                "ShowExtFormatting:<1>" & vbNewLine & _
                "ShowSymbolBar:<1>" & vbNewLine & _
                "ShowStatusBar:<1>" & vbNewLine & _
                "ShowRuler:<1>" & vbNewLine & _
                "WindowState:<2>" & vbNewLine & _
                "Recent1:<>" & vbNewLine & _
                "Recent2:<>" & vbNewLine & _
                "Recent3:<>" & vbNewLine & _
                "Recent4:<>" & vbNewLine & _
                "Recent5:<>" & vbNewLine & _
                "ParseFontTable:<1>"
SetAttr App.Path & "\prefs.prf", vbNormal
SavePrefFile
End Sub

Public Function SavePrefFile()
On Error GoTo 10
Dim FileNum%
FileNum% = FreeFile
Open App.Path & "\prefs.prf" For Output As FileNum%
Print #FileNum%, strPref
Close #FileNum%
Exit Function
10:
If Err.Number = 75 Then
    CustomBox "Could not save preferences to disc because of an access error.", "Make sure the preferences file is not read-only or in use by another application. If this problem persists, hold Ctrl while Hyperwrite starts and choose Reset in the resulting dialog.", vbCritical, "", "", "&OK"
    Exit Function
End If
ErrorTrap
End Function

Public Function DoPrefs(bytOptions As Byte, strOpt As String, Optional strReplace As String = "") As String
'Option: 0 = Load pref only, 1 = Save Pref, Other = Create Pref
    Dim lngOptPos As Long, lngBracketPos As Long, lngOption As Long
    lngOptPos = InStr(1, strPref, strOpt & ":<", vbTextCompare) 'Get the beginning of value
    lngBracketPos = InStr(lngOptPos + 1, strPref, ">", vbBinaryCompare) 'Get the end of value
    If lngOptPos = 0 Or lngBracketPos = 0 Then bytOptions = 2
        lngOption = lngOptPos + Len(strOpt) + 2
    Select Case bytOptions
        Case 0
            DoPrefs = Mid(strPref, lngOption, lngBracketPos - lngOption)
        Case 1
            Dim strLeft As String, strRight As String
            strLeft = Left(strPref, lngOptPos + Len(strOpt) + 1)
            strRight = Mid(strPref, lngBracketPos)
            strPref = strLeft & strReplace & strRight
        Case Else
            If strReplace = "" Then
                strPref = strPref & strOpt & ":<0>" & vbNewLine
            Else
                strPref = strPref & strOpt & ":<" & strReplace & ">" & vbNewLine
            End If
    End Select
End Function
