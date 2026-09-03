Attribute VB_Name = "modTTUM"
'==============================================================================
' modTTUM - Daily TTUM file generator
'
' Builds the bank's fixed-width TTUM upload file from the amounts typed on the
' Entries sheet, for the value date set on the Dashboard.
'
' Record layout (186 characters per record, CRLF separated):
'   1-14    Account number      left justified, space padded
'   15-22   Value date          DDMMYYYY
'   23-45   Amount              23 digits, zero padded, last 2 digits are paise
'   46      Transaction type    D or C
'   47-55   Value date          9 characters: "0" + DDMMYYYY
'   56-81   Filler              26 spaces
'   82-116  Narration           left justified, space padded
'   117-186 Filler              70 spaces
'
' Entry points (Alt+F8, or the buttons on the Dashboard):
'   GenerateTTUM        write the file for the date on the Dashboard
'   PreviewTTUM         show the exact records without writing a file
'   ValidateEntries     check the entries and report problems
'   ResetDateToToday    put today's date back into the value date cell
'   ClearAmounts        blank the amount column, ready for the next day
'   BrowseOutputFolder  pick the output folder
'   OpenOutputFolder    open the output folder in Explorer
'   TTUM_Setup          restore the Dashboard buttons if they are ever missing
'
' Amounts are keyed the way the settlement report shows them: with the Config
' setting on Paise (the default) the last two digits of the figure are the paise,
' so 47400 means 474.00. Set it to Rupees to type 474.00 instead.
'==============================================================================
Option Explicit

' ---- Field widths. Change these only if the bank changes the specification. ----
Private Const LEN_ACCOUNT      As Long = 14
Private Const LEN_DATE1        As Long = 8
Private Const LEN_AMOUNT       As Long = 23
Private Const LEN_TRANTYPE     As Long = 1
Private Const LEN_DATE2        As Long = 9
Private Const LEN_FILLER1      As Long = 26
Private Const LEN_NARRATION    As Long = 35
Private Const LEN_FILLER2      As Long = 70
Private Const LEN_RECORD       As Long = 186

' ---- Sheet and table geometry ----
Private Const SH_DASHBOARD     As String = "Dashboard"
Private Const SH_ENTRIES       As String = "Entries"
Private Const SH_CONFIG        As String = "Config"
Private Const SH_LOG           As String = "Log"
Private Const SH_PREVIEW       As String = "Preview"

Private Const ENTRY_FIRST_ROW  As Long = 5
Private Const ENTRY_LAST_ROW   As Long = 104
Private Const COL_INCLUDE      As Long = 1     ' A
Private Const COL_DESC         As Long = 2     ' B
Private Const COL_ACCOUNT      As Long = 3     ' C
Private Const COL_DRCR         As Long = 4     ' D
Private Const COL_AMOUNT       As Long = 5     ' E
Private Const COL_NARRATION    As Long = 6     ' F

Private Const CLR_ERROR        As Long = 13551615   ' light red
Private Const CLR_NORMAL       As Long = 16777215   ' white
Private Const CLR_INPUT        As Long = 14415871   ' pale yellow, the amount column

' A single validated transaction line, ready to be formatted.
Private Type TTUMEntry
    SourceRow As Long
    Account   As String
    DrCr      As String
    Paise     As Currency
    Narration As String
End Type


'==============================================================================
' Entry points
'==============================================================================

Public Sub GenerateTTUM()
    Dim entries() As TTUMEntry, count As Long
    Dim problems As Collection
    Dim valueDate As Date, fileDate As Date
    Dim folder As String, fileName As String, fullPath As String
    Dim content As String, i As Long
    Dim totalDr As Currency, totalCr As Currency

    On Error GoTo Fail
    Application.ScreenUpdating = False
    ClearRowHighlights

    If Not ReadDates(valueDate, fileDate) Then GoTo Done

    Set problems = New Collection
    count = CollectEntries(valueDate, entries, problems)

    If problems.count > 0 Then
        SetStatus "Not generated - " & problems.count & " problem(s) found.", True
        MsgBox ProblemText(problems), vbExclamation, "TTUM - please fix these first"
        GoTo Done
    End If

    If count = 0 Then
        SetStatus "Not generated - no rows are marked Yes on the Entries sheet.", True
        MsgBox "No entries are included." & vbCrLf & vbCrLf & _
               "Set column A to ""Yes"" for the rows you want in today's file.", _
               vbExclamation, "TTUM - nothing to generate"
        GoTo Done
    End If

    SumTotals entries, count, totalDr, totalCr
    If totalDr <> totalCr Then
        If GetConfigBool("ttEnforceBalance", True) Then
            SetStatus "Not generated - debits and credits do not balance.", True
            MsgBox "The file is out of balance and was not generated." & vbCrLf & vbCrLf & _
                   "Total debit : " & Format$(totalDr, "#,##0.00") & vbCrLf & _
                   "Total credit: " & Format$(totalCr, "#,##0.00") & vbCrLf & _
                   "Difference  : " & Format$(totalDr - totalCr, "#,##0.00") & vbCrLf & vbCrLf & _
                   "Correct the amounts, or set ""Block generation when out of balance"" " & _
                   "to No on the Config sheet.", vbExclamation, "TTUM - out of balance"
            GoTo Done
        Else
            If MsgBox("Debits and credits do not balance." & vbCrLf & vbCrLf & _
                      "Total debit : " & Format$(totalDr, "#,##0.00") & vbCrLf & _
                      "Total credit: " & Format$(totalCr, "#,##0.00") & vbCrLf & _
                      "Difference  : " & Format$(totalDr - totalCr, "#,##0.00") & vbCrLf & vbCrLf & _
                      "Generate the file anyway?", vbExclamation + vbYesNo + vbDefaultButton2, _
                      "TTUM - out of balance") <> vbYes Then
                SetStatus "Cancelled - file is out of balance.", True
                GoTo Done
            End If
        End If
    End If

    folder = ResolveOutputFolder()
    If Not EnsureFolder(folder) Then
        SetStatus "Not generated - output folder could not be created.", True
        MsgBox "The output folder could not be created:" & vbCrLf & vbCrLf & folder, _
               vbCritical, "TTUM - folder error"
        GoTo Done
    End If

    fileName = ResolveFileName(fileDate)
    fullPath = JoinPath(folder, fileName)

    If Len(Dir$(fullPath)) > 0 Then
        If MsgBox("This file already exists:" & vbCrLf & vbCrLf & fullPath & vbCrLf & vbCrLf & _
                  "Replace it?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "TTUM - file already exists") <> vbYes Then
            SetStatus "Cancelled - existing file was kept.", True
            GoTo Done
        End If
    End If

    content = BuildContent(entries, count, valueDate)
    WriteTextFile fullPath, content

    AppendLog valueDate, fileDate, fileName, fullPath, count, totalDr, totalCr
    SetStatus "Generated " & count & " records -> " & fileName & "  (" & Format$(Now, "dd-mmm-yyyy hh:nn") & ")", False

    Application.ScreenUpdating = True
    If MsgBox("TTUM file created." & vbCrLf & vbCrLf & _
              "File     : " & fileName & vbCrLf & _
              "Folder   : " & folder & vbCrLf & _
              "Records  : " & count & vbCrLf & _
              "Value date: " & Format$(valueDate, "dd-mmm-yyyy") & vbCrLf & _
              "Debit    : " & Format$(totalDr, "#,##0.00") & vbCrLf & _
              "Credit   : " & Format$(totalCr, "#,##0.00") & vbCrLf & vbCrLf & _
              "Open the output folder now?", vbInformation + vbYesNo, "TTUM - done") = vbYes Then
        OpenOutputFolder
    End If

Done:
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    SetStatus "Failed - " & Err.Description, True
    MsgBox "The TTUM file could not be generated." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "TTUM - error"
End Sub


Public Sub PreviewTTUM()
    Dim entries() As TTUMEntry, count As Long
    Dim problems As Collection
    Dim valueDate As Date, fileDate As Date
    Dim ws As Worksheet, i As Long, rec As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    ClearRowHighlights

    If Not ReadDates(valueDate, fileDate) Then GoTo Done

    Set problems = New Collection
    count = CollectEntries(valueDate, entries, problems)

    If problems.count > 0 Then
        MsgBox ProblemText(problems), vbExclamation, "TTUM - please fix these first"
        SetStatus "Preview stopped - " & problems.count & " problem(s) found.", True
        GoTo Done
    End If
    If count = 0 Then
        MsgBox "No entries are included. Set column A to ""Yes"" for the rows you want.", _
               vbExclamation, "TTUM - nothing to preview"
        GoTo Done
    End If

    Set ws = GetOrCreateSheet(SH_PREVIEW)
    ws.Cells.Clear
    ws.Range("A1").Value = "PREVIEW - " & ResolveFileName(fileDate) & _
                           "   (value date " & Format$(valueDate, "dd-mmm-yyyy") & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Each line below is exactly " & LEN_RECORD & " characters. " & _
                           "This is a preview only - nothing has been written to disk."
    ws.Range("A2").Font.Italic = True

    ws.Range("A4").Value = "#"
    ws.Range("B4").Value = "Record"
    ws.Range("C4").Value = "Length"
    ws.Range("A4:C4").Font.Bold = True

    For i = 1 To count
        rec = FormatRecord(entries(i), valueDate)
        ws.Cells(4 + i, 1).Value = i
        ws.Cells(4 + i, 2).Value = "'" & rec
        ws.Cells(4 + i, 3).Value = Len(rec)
    Next i

    ws.Columns("B").Font.Name = "Consolas"
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 100
    ws.Columns("C").ColumnWidth = 8
    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A1").Select
    SetStatus "Previewed " & count & " records (no file written).", False

Done:
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Preview failed." & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "TTUM - error"
End Sub


Public Sub ValidateEntries()
    Dim entries() As TTUMEntry, count As Long
    Dim problems As Collection
    Dim valueDate As Date, fileDate As Date
    Dim totalDr As Currency, totalCr As Currency

    ClearRowHighlights
    If Not ReadDates(valueDate, fileDate) Then Exit Sub

    Set problems = New Collection
    count = CollectEntries(valueDate, entries, problems)

    If problems.count > 0 Then
        MsgBox ProblemText(problems), vbExclamation, "TTUM - validation"
        SetStatus "Validation found " & problems.count & " problem(s).", True
        Exit Sub
    End If

    SumTotals entries, count, totalDr, totalCr
    MsgBox "All checks passed." & vbCrLf & vbCrLf & _
           "Rows included: " & count & vbCrLf & _
           "Total debit  : " & Format$(totalDr, "#,##0.00") & vbCrLf & _
           "Total credit : " & Format$(totalCr, "#,##0.00") & vbCrLf & _
           "Difference   : " & Format$(totalDr - totalCr, "#,##0.00"), _
           vbInformation, "TTUM - validation"
    SetStatus "Validation passed for " & count & " rows.", False
End Sub


Public Sub ResetDateToToday()
    On Error Resume Next
    ThisWorkbook.Names("ttValueDate").RefersToRange.Value = Date
    SetStatus "Value date reset to " & Format$(Date, "dd-mmm-yyyy") & ".", False
End Sub


' Blanks the amount column so the sheet is ready for the next day's figures.
Public Sub ClearAmounts()
    Dim ws As Worksheet
    If MsgBox("Clear the amounts in every row, ready for the next day?" & vbCrLf & vbCrLf & _
              "Account numbers, narrations and the Yes/No column are kept.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "TTUM - clear amounts") <> vbYes Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_ENTRIES)
    ws.Range(ws.Cells(ENTRY_FIRST_ROW, COL_AMOUNT), ws.Cells(ENTRY_LAST_ROW, COL_AMOUNT)).ClearContents
    ClearRowHighlights
    SetStatus "Amounts cleared. Enter today's figures on the Entries sheet.", False
End Sub


Public Sub BrowseOutputFolder()
    Dim fd As Object, picked As String
    On Error Resume Next
    Set fd = Application.FileDialog(4)          ' msoFileDialogFolderPicker
    If fd Is Nothing Then
        MsgBox "Type the output folder directly into the Dashboard cell.", vbInformation, "TTUM"
        Exit Sub
    End If
    fd.Title = "Choose the folder the TTUM files should be written to"
    fd.InitialFileName = ResolveOutputFolder() & "\"
    If fd.Show = -1 Then
        picked = fd.SelectedItems(1)
        ThisWorkbook.Names("ttOutputFolder").RefersToRange.Value = picked
        SetStatus "Output folder set to " & picked, False
    End If
End Sub


Public Sub OpenOutputFolder()
    Dim folder As String
    folder = ResolveOutputFolder()
    If Not EnsureFolder(folder) Then
        MsgBox "This folder does not exist and could not be created:" & vbCrLf & vbCrLf & folder, _
               vbExclamation, "TTUM"
        Exit Sub
    End If
    On Error Resume Next
    Shell "explorer.exe """ & folder & """", vbNormalFocus
End Sub


'==============================================================================
' Reading the sheets
'==============================================================================

' Reads the value date and the file-name date from the Dashboard.
Private Function ReadDates(ByRef valueDate As Date, ByRef fileDate As Date) As Boolean
    Dim v As Variant, f As Variant

    On Error GoTo Bad
    v = ThisWorkbook.Names("ttValueDate").RefersToRange.Value
    If Not TryDate(v, valueDate) Then GoTo Bad

    f = ThisWorkbook.Names("ttFileDate").RefersToRange.Value
    If Not TryDate(f, fileDate) Then fileDate = valueDate

    ReadDates = True
    Exit Function
Bad:
    SetStatus "The value date on the Dashboard is not a valid date.", True
    MsgBox "Please enter a valid value date on the Dashboard." & vbCrLf & vbCrLf & _
           "Leave it as today's date, or type the date you want the file to carry.", _
           vbExclamation, "TTUM - date required"
    ReadDates = False
End Function


' Accepts a real date, a date serial number, or a typed date string.
Private Function TryDate(ByVal v As Variant, ByRef result As Date) As Boolean
    On Error GoTo Bad
    If IsEmpty(v) Then GoTo Bad
    If IsDate(v) Then
        result = CDate(v)
    ElseIf IsNumeric(v) Then
        If CDbl(v) < 1 Then GoTo Bad
        result = CDate(CDbl(v))
    Else
        GoTo Bad
    End If
    TryDate = True
    Exit Function
Bad:
    TryDate = False
End Function


' Reads every included row, validates it, and returns how many entries were collected.
' Any problem found is appended to `problems` and the offending cell is highlighted.
Private Function CollectEntries(ByVal valueDate As Date, ByRef entries() As TTUMEntry, _
                                ByRef problems As Collection) As Long
    Dim ws As Worksheet, r As Long, n As Long
    Dim include As String, account As String, drcr As String, narration As String
    Dim rawAmount As Variant, paise As Currency
    Dim isBlank As Boolean, keyedAsPaise As Boolean

    keyedAsPaise = AmountKeyedAsPaise()
    Set ws = ThisWorkbook.Worksheets(SH_ENTRIES)
    ReDim entries(1 To ENTRY_LAST_ROW - ENTRY_FIRST_ROW + 1)
    n = 0

    For r = ENTRY_FIRST_ROW To ENTRY_LAST_ROW
        include = UCase$(Trim$(CStr(ws.Cells(r, COL_INCLUDE).Value)))
        account = CellAsText(ws.Cells(r, COL_ACCOUNT))
        drcr = UCase$(Trim$(CStr(ws.Cells(r, COL_DRCR).Value)))
        narration = Trim$(CStr(ws.Cells(r, COL_NARRATION).Value))
        rawAmount = ws.Cells(r, COL_AMOUNT).Value

        isBlank = (Len(account) = 0 And Len(narration) = 0 And Len(Trim$(CStr(rawAmount))) = 0)
        If isBlank Then GoTo NextRow
        If include <> "YES" And include <> "Y" Then GoTo NextRow

        If Len(account) = 0 Then
            AddProblem problems, r, "account number is empty", ws.Cells(r, COL_ACCOUNT)
            GoTo NextRow
        End If
        If Len(account) > LEN_ACCOUNT Then
            AddProblem problems, r, "account number is " & Len(account) & " characters, the field holds " & _
                       LEN_ACCOUNT, ws.Cells(r, COL_ACCOUNT)
            GoTo NextRow
        End If

        If drcr <> "D" And drcr <> "C" Then
            AddProblem problems, r, "Dr/Cr must be D or C (found """ & _
                       CStr(ws.Cells(r, COL_DRCR).Value) & """)", ws.Cells(r, COL_DRCR)
            GoTo NextRow
        End If

        If Len(Trim$(CStr(rawAmount))) = 0 Then
            AddProblem problems, r, "amount is blank - type today's figure", ws.Cells(r, COL_AMOUNT)
            GoTo NextRow
        End If
        If Not IsNumeric(rawAmount) Then
            AddProblem problems, r, "amount is not a number", ws.Cells(r, COL_AMOUNT)
            GoTo NextRow
        End If
        If keyedAsPaise Then
            If CCur(rawAmount) <> Int(CCur(rawAmount)) Then
                AddProblem problems, r, "amount must be a whole number while Config is set to " & _
                           "Paise, because the last two digits are already the paise", _
                           ws.Cells(r, COL_AMOUNT)
                GoTo NextRow
            End If
        End If
        paise = ToPaise(rawAmount, keyedAsPaise)
        If paise <= 0 Then
            AddProblem problems, r, "amount must be greater than zero", ws.Cells(r, COL_AMOUNT)
            GoTo NextRow
        End If

        narration = SubstituteTokens(narration, valueDate)
        If Len(narration) = 0 Then
            AddProblem problems, r, "narration is empty", ws.Cells(r, COL_NARRATION)
            GoTo NextRow
        End If
        If Len(narration) > LEN_NARRATION Then
            AddProblem problems, r, "narration is " & Len(narration) & " characters, the field holds " & _
                       LEN_NARRATION & " (""" & narration & """)", ws.Cells(r, COL_NARRATION)
            GoTo NextRow
        End If

        n = n + 1
        entries(n).SourceRow = r
        entries(n).Account = account
        entries(n).DrCr = drcr
        entries(n).Paise = paise
        entries(n).Narration = narration
NextRow:
    Next r

    CollectEntries = n
End Function


' Returns the cell's text without letting Excel turn a long account number into
' scientific notation or drop its leading zeros.
Private Function CellAsText(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsEmpty(v) Then
        CellAsText = ""
    ElseIf VarType(v) = vbDouble Or VarType(v) = vbSingle Or VarType(v) = vbLong Or _
           VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        CellAsText = Trim$(Format$(v, "0"))
    Else
        CellAsText = Trim$(CStr(v))
    End If
End Function


Private Sub AddProblem(ByRef problems As Collection, ByVal r As Long, ByVal msg As String, _
                       ByVal target As Range)
    problems.Add "Row " & r & ": " & msg
    On Error Resume Next
    target.Interior.Color = CLR_ERROR
End Sub


Private Function ProblemText(ByVal problems As Collection) As String
    Dim s As String, i As Long, shown As Long
    shown = problems.count
    If shown > 15 Then shown = 15
    s = "The file was not generated. Please fix the highlighted cells on the Entries sheet:" & _
        vbCrLf & vbCrLf
    For i = 1 To shown
        s = s & "  - " & problems(i) & vbCrLf
    Next i
    If problems.count > shown Then
        s = s & "  ... and " & (problems.count - shown) & " more." & vbCrLf
    End If
    ProblemText = s
End Function


' Puts the entry grid back to its normal colours, keeping the amount column
' shaded as an input field.
Private Sub ClearRowHighlights()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_ENTRIES)
    If ws Is Nothing Then Exit Sub
    ws.Range(ws.Cells(ENTRY_FIRST_ROW, COL_INCLUDE), _
             ws.Cells(ENTRY_LAST_ROW, COL_NARRATION)).Interior.Color = CLR_NORMAL
    ws.Range(ws.Cells(ENTRY_FIRST_ROW, COL_AMOUNT), _
             ws.Cells(ENTRY_LAST_ROW, COL_AMOUNT)).Interior.Color = CLR_INPUT
End Sub


Private Sub SumTotals(ByRef entries() As TTUMEntry, ByVal count As Long, _
                      ByRef totalDr As Currency, ByRef totalCr As Currency)
    Dim i As Long
    totalDr = 0: totalCr = 0
    For i = 1 To count
        If entries(i).DrCr = "D" Then
            totalDr = totalDr + entries(i).Paise / 100
        Else
            totalCr = totalCr + entries(i).Paise / 100
        End If
    Next i
End Sub


'==============================================================================
' Formatting the file
'==============================================================================

Private Function BuildContent(ByRef entries() As TTUMEntry, ByVal count As Long, _
                              ByVal valueDate As Date) As String
    Dim parts() As String, i As Long, s As String
    ReDim parts(1 To count)
    For i = 1 To count
        parts(i) = FormatRecord(entries(i), valueDate)
    Next i
    s = Join(parts, vbCrLf)
    If GetConfigBool("ttTrailingNewline", False) Then s = s & vbCrLf
    BuildContent = s
End Function


' Lays one entry out across the 186 character record.
Private Function FormatRecord(ByRef e As TTUMEntry, ByVal valueDate As Date) As String
    Dim ddmmyyyy As String, rec As String
    ddmmyyyy = Format$(valueDate, "ddmmyyyy")

    rec = PadRight(e.Account, LEN_ACCOUNT) & _
          ddmmyyyy & _
          PadZeros(PaiseToDigits(e.Paise), LEN_AMOUNT) & _
          e.DrCr & _
          "0" & ddmmyyyy & _
          Space$(LEN_FILLER1) & _
          PadRight(e.Narration, LEN_NARRATION) & _
          Space$(LEN_FILLER2)

    If Len(rec) <> LEN_RECORD Then
        Err.Raise vbObjectError + 513, "FormatRecord", _
                  "Row " & e.SourceRow & " produced a " & Len(rec) & " character record; " & _
                  LEN_RECORD & " were expected."
    End If
    FormatRecord = rec
End Function


' Turns what was keyed into the Amount column into whole paise.
'
' With the Config setting on Paise - the default - the figure is already the one
' the settlement report shows, with the last two digits being the paise, so it is
' used as it stands. On Rupees it is a rupee amount and gets multiplied by 100.
'
' Currency arithmetic is exact to four decimals, so this avoids the rounding drift
' a Double would introduce on large settlement values.
Private Function ToPaise(ByVal amount As Variant, ByVal keyedAsPaise As Boolean) As Currency
    Dim c As Currency
    If keyedAsPaise Then
        c = CCur(amount)
    Else
        c = CCur(amount) * 100
    End If
    ToPaise = Int(c + CCur(0.5))
End Function


' True when the Amount column holds the figure exactly as the report shows it.
Private Function AmountKeyedAsPaise() As Boolean
    AmountKeyedAsPaise = (UCase$(GetConfigText("ttAmountUnit", "Paise")) <> "RUPEES")
End Function


Private Function PaiseToDigits(ByVal paise As Currency) As String
    PaiseToDigits = Format$(paise, "0")
End Function


Private Function PadRight(ByVal s As String, ByVal width As Long) As String
    If Len(s) >= width Then
        PadRight = Left$(s, width)
    Else
        PadRight = s & Space$(width - Len(s))
    End If
End Function


Private Function PadZeros(ByVal s As String, ByVal width As Long) As String
    If Len(s) > width Then
        Err.Raise vbObjectError + 514, "PadZeros", _
                  "The amount " & s & " needs " & Len(s) & " digits but the field holds " & width & "."
    End If
    PadZeros = String$(width - Len(s), "0") & s
End Function


' Replaces the date placeholders a narration template may contain.
' Longest tokens first so {DDMMYYYY} is not eaten by {DDMMYY}.
Public Function SubstituteTokens(ByVal template As String, ByVal d As Date) As String
    Dim s As String
    s = template
    s = ReplaceToken(s, "{DDMMYYYY}", Format$(d, "ddmmyyyy"))
    s = ReplaceToken(s, "{DDMMYY}", Format$(d, "ddmmyy"))
    s = ReplaceToken(s, "{YYYY}", Format$(d, "yyyy"))
    s = ReplaceToken(s, "{DD}", Format$(d, "dd"))
    s = ReplaceToken(s, "{MM}", Format$(d, "mm"))
    s = ReplaceToken(s, "{YY}", Format$(d, "yy"))
    SubstituteTokens = s
End Function


Private Function ReplaceToken(ByVal s As String, ByVal token As String, ByVal value As String) As String
    ReplaceToken = Replace(s, token, value, 1, -1, vbTextCompare)
End Function


'==============================================================================
' Files, config and logging
'==============================================================================

' Writes the text as plain single-byte characters with no trailing byte the
' bank did not ask for.
Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim f As Integer, b() As Byte
    If Len(Dir$(path)) > 0 Then Kill path
    b = StrConv(content, vbFromUnicode)
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f
End Sub


Private Function ResolveOutputFolder() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(ThisWorkbook.Names("ttOutputFolder").RefersToRange.Value))
    On Error GoTo 0
    If Len(s) = 0 Then s = JoinPath(ThisWorkbook.path, "TTUM_Output")
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    ResolveOutputFolder = s
End Function


Private Function ResolveFileName(ByVal fileDate As Date) As String
    Dim pattern As String
    pattern = GetConfigText("ttFileNamePattern", "PROPELG_TTUM_{DDMMYYYY}.txt")
    ResolveFileName = SubstituteTokens(pattern, fileDate)
End Function


Private Function JoinPath(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Then
        JoinPath = a & b
    Else
        JoinPath = a & "\" & b
    End If
End Function


' Creates the folder, and any missing parent, returning False if that is not possible.
Private Function EnsureFolder(ByVal path As String) As Boolean
    Dim parent As String
    On Error GoTo Bad
    If Len(Dir$(path, vbDirectory)) > 0 Then
        EnsureFolder = True
        Exit Function
    End If
    parent = Left$(path, InStrRev(path, "\") - 1)
    If Len(parent) > 2 Then
        If Not EnsureFolder(parent) Then GoTo Bad
    End If
    MkDir path
    EnsureFolder = True
    Exit Function
Bad:
    EnsureFolder = False
End Function


Private Function GetConfigText(ByVal nameRef As String, ByVal fallback As String) As String
    Dim s As String
    On Error GoTo UseFallback
    s = Trim$(CStr(ThisWorkbook.Names(nameRef).RefersToRange.Value))
    If Len(s) = 0 Then GoTo UseFallback
    GetConfigText = s
    Exit Function
UseFallback:
    GetConfigText = fallback
End Function


Private Function GetConfigBool(ByVal nameRef As String, ByVal fallback As Boolean) As Boolean
    Dim s As String
    s = UCase$(GetConfigText(nameRef, ""))
    Select Case s
        Case "YES", "Y", "TRUE", "1": GetConfigBool = True
        Case "NO", "N", "FALSE", "0": GetConfigBool = False
        Case Else:                    GetConfigBool = fallback
    End Select
End Function


Private Sub SetStatus(ByVal msg As String, ByVal isProblem As Boolean)
    Dim c As Range
    On Error Resume Next
    Set c = ThisWorkbook.Names("ttStatus").RefersToRange
    If c Is Nothing Then Exit Sub
    c.Value = msg
    If isProblem Then
        c.Font.Color = RGB(170, 0, 0)
    Else
        c.Font.Color = RGB(0, 110, 60)
    End If
End Sub


Private Sub AppendLog(ByVal valueDate As Date, ByVal fileDate As Date, ByVal fileName As String, _
                      ByVal fullPath As String, ByVal count As Long, _
                      ByVal totalDr As Currency, ByVal totalCr As Currency)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.count, 1).End(-4162).Row + 1     ' xlUp
    If r < 4 Then r = 4
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Cells(r, 2).Value = valueDate
    ws.Cells(r, 2).NumberFormat = "dd-mmm-yyyy"
    ws.Cells(r, 3).Value = fileDate
    ws.Cells(r, 3).NumberFormat = "dd-mmm-yyyy"
    ws.Cells(r, 4).Value = fileName
    ws.Cells(r, 5).Value = count
    ws.Cells(r, 6).Value = CDbl(totalDr)
    ws.Cells(r, 7).Value = CDbl(totalCr)
    ws.Cells(r, 8).Value = IIf(totalDr = totalCr, "Balanced", "OUT OF BALANCE")
    ws.Cells(r, 9).Value = Application.UserName
    ws.Cells(r, 10).Value = fullPath
    ws.Cells(r, 6).NumberFormat = "#,##0.00"
    ws.Cells(r, 7).NumberFormat = "#,##0.00"
End Sub


'==============================================================================
' Dashboard buttons
'==============================================================================

' Rebuilds the Dashboard buttons. Runs on open, and can be run by hand from
' Alt+F8 if the buttons are ever deleted.
Public Sub TTUM_Setup()
    Dim ws As Worksheet, shp As Shape, i As Long, present As Long
    Dim labels As Variant, macros As Variant
    Dim topPos As Double

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DASHBOARD)
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub

    FillDefaultOutputFolder

    labels = Array("Generate TTUM File", "Preview Records", "Validate Entries", _
                   "Reset Date to Today", "Clear Amounts", "Choose Output Folder", _
                   "Open Output Folder")
    macros = Array("GenerateTTUM", "PreviewTTUM", "ValidateEntries", _
                   "ResetDateToToday", "ClearAmounts", "BrowseOutputFolder", _
                   "OpenOutputFolder")

    ' The workbook already carries its buttons. When they are there, just make
    ' sure each one still points at its macro - that is what repairs a button
    ' that looks right but does nothing. Only build new ones if they are gone.
    present = 0
    For i = 1 To ws.Shapes.count
        If Left$(ws.Shapes(i).Name, 5) = "btnTT" Then present = present + 1
    Next i
    If present >= UBound(labels) - LBound(labels) + 1 Then
        For i = LBound(labels) To UBound(labels)
            On Error Resume Next
            ws.Shapes("btnTT" & i).OnAction = "modTTUM." & macros(i)
            On Error GoTo Fail
        Next i
        Exit Sub
    End If

    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "btnTT" Then ws.Shapes(i).Delete
    Next i

    topPos = ws.Range("H5").Top
    For i = LBound(labels) To UBound(labels)
        Set shp = ws.Shapes.AddFormControl(0, ws.Range("H5").Left, _
                                           topPos + i * 38, 165, 30)   ' xlButtonControl
        shp.Name = "btnTT" & i
        shp.OnAction = "modTTUM." & macros(i)
        shp.TextFrame.Characters.Text = labels(i)
        shp.TextFrame.Characters.Font.Size = 10
        If i = 0 Then shp.TextFrame.Characters.Font.Bold = True
    Next i
    Exit Sub
Fail:
    MsgBox "The Dashboard buttons could not be created." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "You can still run every command from Alt+F8.", vbExclamation, "TTUM - setup"
End Sub


' Puts a real, editable path in the output folder cell the first time the
' workbook is opened, so the folder is obvious rather than an empty box.
Private Sub FillDefaultOutputFolder()
    Dim c As Range
    On Error Resume Next
    Set c = ThisWorkbook.Names("ttOutputFolder").RefersToRange
    If c Is Nothing Then Exit Sub
    If Len(Trim$(CStr(c.Value))) > 0 Then Exit Sub
    If Len(ThisWorkbook.path) = 0 Then Exit Sub
    c.Value = JoinPath(ThisWorkbook.path, "TTUM_Output")
End Sub


Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
