Attribute VB_Name = "Module1"
' ------------------------------------------------------------------------------
' Subroutine: ResetSettings
' Description: Restores common Excel settings to default for troubleshooting.
' ------------------------------------------------------------------------------

Sub ResetSettings()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = False
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.UsedRange ' Refresh UsedRange reference
End Sub


' ------------------------------------------------------------------------------
' Subroutine: TaxonomyKeyParse
' Description: Converts client taxonomy strings (formatted as KEY~Value_) into a
' structured table, separating keys into columns.
' ------------------------------------------------------------------------------

Sub TaxonomyKeyParse()
    If Selection.Cells.Count = 0 Then
        MsgBox "Select one or more cells with taxonomy strings.", vbExclamation
        Exit Sub
    End If

    Dim cell As Range, segs As Variant, kv() As String
    Dim dictHeaders As Object: Set dictHeaders = CreateObject("Scripting.Dictionary")
    Dim dataRows() As Object
    Dim i As Long, actualRowCount As Long

    ReDim dataRows(1 To Selection.Cells.Count)
    actualRowCount = 0

    ' Parse each selected cell
    For Each cell In Selection.Cells
        If Trim(cell.Value) <> "" Then
            actualRowCount = actualRowCount + 1
            Set dataRows(actualRowCount) = CreateObject("Scripting.Dictionary")
            segs = Split(cell.Value, "_")
            For i = 0 To UBound(segs)
                If InStr(segs(i), "~") > 0 Then
                    kv = Split(segs(i), "~", 2)
                    If Not dictHeaders.Exists(kv(0)) Then dictHeaders.Add kv(0), dictHeaders.Count + 1
                    dataRows(actualRowCount).Add kv(0), kv(1)
                End If
            Next i
        End If
    Next cell

    If actualRowCount = 0 Then
        MsgBox "No valid taxonomy strings found.", vbExclamation
        Exit Sub
    End If

    ' Generate new worksheet for parsed output
    Dim wsName As String: wsName = "Key Parse"
    Dim N As Long: N = 0
    Do While SheetExists(wsName)
        N = N + 1
        wsName = "Key Parse-" & N
    Loop
    Dim ws As Worksheet: Set ws = Worksheets.Add
    ws.Name = wsName

    ' Write headers
    Dim headerKeys() As Variant: headerKeys = dictHeaders.Keys
    For i = 0 To UBound(headerKeys)
        ws.Cells(1, i + 1).Value = headerKeys(i)
    Next i

    ' Write each parsed data row
    Dim r As Long, c As Long
    For r = 1 To actualRowCount
        For c = 0 To UBound(headerKeys)
            If dataRows(r).Exists(headerKeys(c)) Then
                ws.Cells(r + 1, c + 1).Value = dataRows(r)(headerKeys(c))
            End If
        Next c
    Next r

    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    ws.Activate
End Sub

' Check if a worksheet with the given name exists
Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ------------------------------------------------------------------------------
' Subroutine: MergeExcelSheets
' Description: Merges data (excluding headers after first file) from selected
' workbooks and worksheets into one master workbook with formatting.
' ------------------------------------------------------------------------------

Sub MergeExcelSheets()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim fd As FileDialog, f As Variant, wbSrc As Workbook, wbDest As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet, hdrDone As Boolean
    Dim lastDestRow As Long, rngData As Range
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = True
    If fd.Show = 0 Then GoTo CleanExit

    Set wbDest = Workbooks.Add
    Set wsDest = wbDest.Sheets(1)
    wsDest.Name = "Merged Data"
    
    For Each f In fd.SelectedItems
        Set wbSrc = Workbooks.Open(f)
        If wbSrc.ProtectStructure Then wbSrc.Windows(1).WindowState = xlNormal
        
        For Each wsSrc In wbSrc.Worksheets
            If LCase(wsSrc.Name) <> "help" Then
                With wsSrc.UsedRange
                    If Not hdrDone Then
                        .Rows(1).Copy wsDest.Range("A1")
                        hdrDone = True
                    End If
                    If .Rows.Count > 1 Then
                        Set rngData = .Offset(1).Resize(.Rows.Count - 1)
                        lastDestRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
                        rngData.Copy wsDest.Cells(lastDestRow, 1)
                    End If
                End With
            End If
        Next wsSrc
        wbSrc.Close False
    Next f

    ' Apply header formatting
    With wsDest
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(242, 242, 242)
        .Rows(1).HorizontalAlignment = xlCenter
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).AutoFilter
        .Columns.AutoFit
    End With

    With ActiveWindow
        .SplitRow = 1
        .SplitColumn = 0
        .FreezePanes = True
    End With

    Call QuickSave

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


' ------------------------------------------------------------------------------
' Subroutine: FormatCSV
' Description: Formats CSV-style sheet with styled headers, alignment, conditional
'              number/date formatting, and controlled auto-width. Saves to file.
' ------------------------------------------------------------------------------
Sub FormatCSV()
    Dim hdr As Range, c As Range, colData As Range
    Set hdr = Range("A1", Cells(1, Cells(1, Columns.Count).End(xlToLeft).Column))    'Header range from A1 to last used column

    If Not hdr.Worksheet.AutoFilterMode Then hdr.AutoFilter                         'Enable autofilter if not already on

    With hdr                                                                         'Header formatting
        .Interior.Color = RGB(242, 242, 242)                                         'Grey fill
        .Font.Bold = True                                                            'Bold font
        .HorizontalAlignment = xlCenter                                              'Center text horizontally
        .VerticalAlignment = xlCenter                                                'Center text vertically
        .Rows.RowHeight = 20                                                         'Set header row height
    End With

    Rows("2:" & Cells(Rows.Count, 1).End(xlUp).Row).RowHeight = 15                   'Set all data row heights to 15

    With Range("A2", Cells(Cells(Rows.Count, 1).End(xlUp).Row, hdr.Columns.Count))   'All data cells range
        .HorizontalAlignment = xlLeft                                                'Align data rows left
    End With

    For Each c In hdr
        Dim headerText As String
        headerText = UCase(c.Value)
        Set colData = Range(c.Offset(1), Cells(Rows.Count, c.Column).End(xlUp))      'Data range below header

        If headerText Like "*DATE*" Or headerText Like "*DAY*" Or _
           headerText Like "*MONTH*" Or headerText Like "*YEAR*" Or _
           headerText Like "*CALENDAR*" Then

            colData.NumberFormat = "m/d/yyyy"                                        'Short date format
            'colData.IndentLevel = 1                                                 'Slightly indent for readability
            c.EntireColumn.AutoFit                                                   'Auto-size column based on content
            c.ColumnWidth = c.ColumnWidth + 2                                        'Buffer for zoom/display
        ElseIf headerText Like "*ID*" Or headerText Like "*KEY*" Or headerText Like "*CODE*" Then
            colData.NumberFormat = "0"                                               'No decimals for ID/Key/Code
            c.EntireColumn.AutoFit                                                   'Auto-size column based on content
            c.ColumnWidth = c.ColumnWidth + 2                                        'Buffer for zoom/display
        ElseIf IsNumeric(Application.WorksheetFunction.Index(colData.Value, 1)) Then
            colData.NumberFormat = "#,##0"                                           'Comma format for numbers
            c.ColumnWidth = Len(c.Value) + 2                                         'Use header length only
        Else
            c.ColumnWidth = Len(c.Value) + 2                                         'Default: width based on header only
        End If
    Next c

    With ActiveWindow                                                                'Freeze top row
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

    Call QuickSave                                                                   'Save as new file
End Sub

' ------------------------------------------------------------------------------
' Subroutine: QuickSave
' Description: Saves active workbook as .xlsx in Downloads with date appended.
'              Adds _1, _2, etc. suffixes if file already exists.
' ------------------------------------------------------------------------------
Sub QuickSave()
    On Error GoTo eh1

    Dim savePath As String, baseName As String, fullPath As String
    Dim coreName As String, i As Long

    savePath = Environ$("USERPROFILE") & "\Downloads\"

    ' Strip extension if present
    coreName = ActiveWorkbook.Name
    If InStr(coreName, ".") > 0 Then coreName = Left(coreName, InStrRev(coreName, ".") - 1)

    ' Remove existing _# suffix if present
    If coreName Like "*_*" Then
        Dim underscorePos As Long
        underscorePos = InStrRev(coreName, "_")
        If IsNumeric(Mid(coreName, underscorePos + 1)) Then
            coreName = Left(coreName, underscorePos - 1)
        End If
    End If

    ' Remove trailing date if present
    If Len(coreName) > 11 And Right(coreName, 10) Like "##.##.####" Then
        If Mid(coreName, Len(coreName) - 10, 1) = " " Then
            coreName = Left(coreName, Len(coreName) - 11)
        End If
    End If

    baseName = coreName & " " & Format(Date, "MM.DD.YYYY")
    fullPath = savePath & baseName & ".xlsx"
    i = 1

    Do While FileExists(fullPath)
        fullPath = savePath & baseName & "_" & i & ".xlsx"
        i = i + 1
    Loop

    ActiveWorkbook.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Exit Sub

eh1:
    Dim fName As Variant
    MsgBox "Auto-save failed (file open, permission, or path issue). Please choose where to save.", vbInformation
    fName = Application.GetSaveAsFilename( _
                InitialFileName:=savePath & coreName & ".xlsx", _
                FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If fName <> False Then ActiveWorkbook.SaveAs Filename:=fName, FileFormat:=xlOpenXMLWorkbook
    Exit Sub
End Sub


' ------------------------------------------------------------------------------
' Subroutine: UnprotectSheets
' Description: Opens any Excel files stuck in Protected View so other macros
' can operate without interruption.
' ------------------------------------------------------------------------------

Sub UnprotectSheets()
    Dim wb As Workbook
    Dim wbPV As ProtectedViewWindow

    If Application.ProtectedViewWindows.Count > 0 Then
        For Each wbPV In Application.ProtectedViewWindows
            wbPV.Activate
            Set wb = wbPV.Edit()
        Next wbPV
    End If

    Set wb = Nothing
End Sub

' ------------------------------------------------------------------------------
' Subroutine: ReplaceTildes
' Description: Replaces common instances of double tildes (~~) in the active range.
' for quick Excel lookups and find dialogues
' ------------------------------------------------------------------------------

Sub ReplaceTildes()
    ' Ensure something is selected before running replace
    If Not TypeOf Selection Is Range Then
        MsgBox "Please select a valid range to perform replace.", vbExclamation
        Exit Sub
    End If

    ' Replace double tildes with themselves (appears to be a placeholder action)
    Selection.Replace What:="~~", Replacement:="~~", LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
End Sub

' ------------------------------------------------------------------------------
' Function: FileExists
' Description: Returns True when the supplied file path refers to an existing file.
' ------------------------------------------------------------------------------
Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function


' ------------------------------------------------------------------------------
' Function: FileInUse
' Description: Returns True when the specified file is currently open/locked by
'              another process or user.
' ------------------------------------------------------------------------------
Function FileInUse(filePath As String) As Boolean
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #ff
    If Err.Number <> 0 Then
        FileInUse = True
    Else
        FileInUse = False
        Close #ff
    End If
    On Error GoTo 0
End Function

