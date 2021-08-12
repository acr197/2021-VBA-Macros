Sub format_extender()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    
    'If Sheet isn't 'traffic workbook', bring user to End Sub
    If Not ActiveSheet.Name = "Traffic Workbook" Then
        MsgBox ("Select the Traffic Workbook tab before running")
        GoTo EndMacro
        Else
    End If
        'If Cell G12 isn't status or P12 isnt verification, someone likely inserted a column before, which may make affect where conditional formatting is placed
        'This lets the user decide if they want to move the column, or continue trying to fix formatting at their own risk
        Dim Result As Integer
        If Not Range("G12").Text = "Status" Or Not Range("P12").Text = "Verification" Then
            Result = MsgBox("It was detected a column may have been added or header was renamed prior to the URL columns, which may impact this macros ability to refresh the formatting correctly." & vbNewLine & vbNewLine & "Would you still like to run the macro?", vbYesNo)
            If Result = vbYes Then
            GoTo RunMacro
            Else
            MsgBox "No changes were made." & vbNewLine & vbNewLine & "To fix this error, try one of the following" & vbNewLine & _
            Chr(149) & "Ensure any new added columns are as far to the right as possible" & vbNewLine & _
            Chr(149) & "Ensure no headers that were already built have been renamed"
            GoTo EndMacro
        End If
    End If
    
RunMacro:
        
        'Clear All Formatting
        Range("B1:XFD1,F4:XFD11,AG2:XFD3,AG12:XFD12,13:1048576").Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
                 :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        Cells.FormatConditions.Delete
        Selection.ClearHyperlinks
        Rows("13:13").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearFormats
        
        'Store the amount of columns to format based on the legnth of Row 12
        Dim lngLastColumn  As Long
        Range("A12").Select
        lngLastColumn = Cells(12, Columns.Count).End(xlToLeft).Column - 1
        Range("A13").Resize(, Selection.Columns.Count + lngLastColumn).Select
        
        'Conditional Formatting: Status Row 13
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$G13=""REVIEW"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 15853276
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""IN PROCESS"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 14994616
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""PAUSED"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16755669
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""HOLD"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 12632256
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""PAUSE CREATIVE"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 11381759
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""PAUSE PLACEMENT"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 7316989
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""UPDATE"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 11657469
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G13=""NEW"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 11336585
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Conditional Format: End Dates (Placement and Asset)
        Rows("12:12").Select
        Selection.Find(What:="End Date", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.Resize(, 2).Select
        Selection.FormatConditions.Add Type:=xlTimePeriod, DateOperator:= _
                                       xlNextMonth
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        Selection.FormatConditions.Add Type:=xlTimePeriod, DateOperator:= _
                                       xlThisMonth
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Range("C13:E13").Select
        Selection.NumberFormat = "m/d/yyyy"
        
        'Conditional Format: If Placement information is filled out, Weight, AdChoices, Survey, Verification, and Click Thru 1 must be filled out
        Rows("12:12").Select
        Selection.Find(What:="Weight", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Resize(, 5).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$B13="""""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Conditional Format: If the dimension listed under "dimensions" is not in creative or placement name, highlight red, excluding 1x1s
        Rows("12:12").Select
        Selection.Find(What:="Dimension", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=IF(OR($F13=""1x1"",$F13=""1 x 1""),"""",OR(NOT(ISNUMBER(SEARCH(SUBSTITUTE(F13,"" "",""""),B13))),NOT(ISNUMBER(SEARCH(SUBSTITUTE(F13,"" "",""""),J13)))))"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = 0
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Conditional Format: If Verification is Blocking and Dimension is 1x1, highlight red
        Rows("12:12").Select
        Selection.Find(What:="Verification", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=AND($P13=""Blocking & Monitoring"",OR($F13=""1x1"",$F13=""1 x 1""))"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Conditional Format: If AdChoices is Blocking and Dimension is 1x1, highlight red
        Rows("12:12").Select
        Selection.Find(What:="AdChoices", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=AND(OR($N13=""Upper Right"",$N13=""Upper Left"",$N13=""Lower Right"",$N13=""Lower Left""),OR($F13=""1x1"",$F13=""1 x 1""))"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
            
        'Conditional Format: If URL has space in it, highlight red
        Rows("12:12").Select
        Selection.Find(What:="Click-Thru URL 1", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=FIND("" "",$Q13,1)"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Conditional Format: If URL is missing http, highlight red
        Rows("12:12").Select
        Selection.Find(What:="Click-Thru URL 1", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=NOT(ISNUMBER(SEARCH(""http"",$Q13,1)))"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Validation: Status
        Rows("12:12").Select
        Selection.Find(What:="Status", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:= _
                 "NEW,UPDATE,PAUSE PLACEMENT,PAUSE CREATIVE,HOLD,PAUSED,LIVE,REVIEW,IN PROGRESS"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        'Validation: AdChoices
        Rows("12:12").Select
        Selection.Find(What:="AdChoices", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:= _
                 "Upper Right,Upper Left,Lower Right,Lower Left,Implemented With Pub/DSP, Custom (See Notes), None"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        'Validation: Survey Pixel
        Rows("12:12").Select
        Selection.Find(What:="Survey", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:="Yes,No"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        'Validation: Verification
        Rows("12:12").Select
        Selection.Find(What:="Verification", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:="Blocking & Monitoring, Monitoring Only, None"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        'Validation: Weight % (uses OR/AND/SEARCH formula to only allow number from 0-100, OR the term 'even')
        Rows("12:12").Select
        Selection.Find(What:="Weight", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:= _
                 "=OR(AND(M13>=0,M13<=100),ISNUMBER(SEARCH(""Even"",M13)),ISNUMBER(SEARCH(""Optimized"",M13)))"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "Error"
            .InputMessage = ""
            .ErrorMessage = _
                            "Please enter a Weight value between 0-100, Or put ""Even"""
            .ShowInput = False
            .ShowError = True
        End With
        
        'Convert Start/End dates from General to Date format
        Range(Range("C13").End(xlDown), Cells(Range("B13").End(xlUp).Row + 11, "E")).Select
        Selection.NumberFormat = "m/d/yyyy"
        With Selection
            .HorizontalAlignment = xlLeft
        End With
        
        'Reset Column Width
        Range("A1").Select
        Cells.Find(What:="Site Name", After:=ActiveCell, LookIn:=xlFormulas, _
                                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 26
        Cells.Find(What:="Associated Placement", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 100
        Cells.Find(What:="Start Date", After:=ActiveCell, LookIn:=xlFormulas, _
                                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 17
        Cells.Find(What:="End Date", After:=ActiveCell, LookIn:=xlFormulas, _
                                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 17
        Cells.Find(What:="Asset Expiration Date", After:=ActiveCell, LookIn:=xlFormulas, _
                                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 25
        Cells.Find(What:="Dimensions", After:=ActiveCell, LookIn:=xlFormulas, _
                                       LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                       MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 17
        Cells.Find(What:="Status", After:=ActiveCell, LookIn:=xlFormulas, _
                                   LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                   MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 17
        Cells.Find(What:="Daypart", After:=ActiveCell, LookIn:=xlFormulas, _
                                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 30
        Cells.Find(What:="Notes", After:=ActiveCell, LookIn:=xlFormulas, _
                                  LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                  MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 50
        Cells.Find(What:="Primary Creative Name", After:=ActiveCell, LookIn:=xlFormulas, _
                                  LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                  MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 30
        Cells.Find(What:="Static Backup URL", After:=ActiveCell, LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                 MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 30
        Cells.Find(What:="Weight %", After:=ActiveCell, LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                 MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 15
        Cells.Find(What:="AdChoices", After:=ActiveCell, LookIn:=xlFormulas, _
                                      LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                      MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 20
        Cells.Find(What:="Survey Pixel", After:=ActiveCell, LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                 MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 20
        Cells.Find(What:="Verification", After:=ActiveCell, LookIn:=xlFormulas, _
                                         LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                         MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 25
        Cells.Find(What:="Click-Thru URL 1", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 40
        Cells.Find(What:="Click-Thru URL 2", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 40
        Cells.Find(What:="Click-Thru URL 3", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 40
        Cells.Find(What:="Click-Thru URL 4", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 40
        Cells.Find(What:="Click-Thru URL 5", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 20
        Cells.Find(What:="Click-Thru URL 6", After:=ActiveCell, LookIn:=xlFormulas, _
                                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                     MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 20
        Cells.Find(What:="Tagged URL", After:=ActiveCell, LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                 MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.ColumnWidth = 175
        
        'Copy/Paste all formatting/validation built in Row 13 down to bottom of TWB
        Dim lngLastRow2 As Long
        lngLastRow2 = Sheets("Traffic Workbook").Cells(Rows.Count, "A").End(xlUp).Row
        
        Range("A13").Resize(, Selection.Columns.Count + lngLastColumn).Select
        Selection.Copy
        Range(Selection, Selection.End(xlDown)).Select
        
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                               SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
                               SkipBlanks:=False, Transpose:=False
        
        With Selection.Font
            .Name = "Calibri"
            .Size = 11
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        ''''Defaults
        
        'Default Conditional Formatting: Status Row 4
        Range("F4:Q4").Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""REVIEW"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 15853276
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""IN PROCESS"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 14994616
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""PAUSED"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16755669
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""HOLD"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 12632256
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""PAUSE CREATIVE"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 11381759
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""PAUSE PLACEMENT"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 7316989
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""UPDATE"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 11657469
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=$G4=""NEW"""
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 11336585
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'Default Validation: Status
        Rows("3:3").Select
        Selection.Find(What:="Status", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:= _
                 "NEW,UPDATE,PAUSE PLACEMENT,PAUSE CREATIVE,HOLD,PAUSED,LIVE,REVIEW,IN PROGRESS"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        'Default Conditional Format: Asset Expiration Date
        Rows("3:3").Select
        Selection.Find(What:="Asset Expiration", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlTimePeriod, DateOperator:= _
                                       xlNextMonth
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        Selection.FormatConditions.Add Type:=xlTimePeriod, DateOperator:= _
                                       xlThisMonth
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        ActiveCell.Resize(8, 1).Select
        Selection.NumberFormat = "m/d/yyyy"
        With Selection
            .HorizontalAlignment = xlLeft
        End With
        
        'default: Conditional Format If URL has space in it, highlight red
        Rows("3:3").Select
        Selection.Find(What:="Click-through URL", After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=FIND("" "",$I4,1)"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        'default conditional Format If URL is missing http, highlight red _
         removed because all default cells stay red until filled out, appearing like there are errors where there are not _
         uncomment the below if the highlighting is still useful to you.
'        Rows("3:3").Select
'        Selection.Find(What:="Click-through URL", After:=ActiveCell, LookIn:=xlFormulas, _
'            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
'            MatchCase:=False, SearchFormat:=False).Activate
'        ActiveCell.Offset(1, 0).Select
'        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
'            "=NOT(ISNUMBER(SEARCH(""http"",$I4,1)))"
'        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'        With Selection.FormatConditions(1).Interior
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorAccent2
'            .TintAndShade = 0.799981688894314
'        End With
'        Selection.FormatConditions(1).StopIfTrue = False
        
        'Default: Fill in section grey with white border
        Range("F4:W4").Select
        With Selection.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 14277081
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlThick
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlThick
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        
        'Copy Formatting from Row 4 down in Default section
        Range("F4:W4").Copy
        Range("F4:W11").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                               SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
                               SkipBlanks:=False, Transpose:=False
                               
        'Default merge the cells
        Range("K3:L3,K4:L4,K5:L5,K6:L6,K7:L7,K8:L8,K9:L9,K10:L10,K11:L11,M3:Q3,M4:Q4,M5:Q5,M6:Q6,M7:Q7,M8:Q8,M9:Q9,M10:Q10,M11:Q11,U3:V3,U4:V4,U5:V5,U6:V6,U7:V7,U8:V8,U9:V9,U10:V10,U11:V11").Select
        Range("K3:L3,K4:L4,K5:L5,K6:L6,K7:L7,K8:L8,K9:L9,K10:L10,K11:L11,M3:Q3,M4:Q4,M5:Q5,M6:Q6,M7:Q7,M8:Q8,M9:Q9,M10:Q10,M11:Q11,U3:V3,U4:V4,U5:V5,U6:V6,U7:V7,U8:V8,U9:V9,U10:V10,U11:V11").UnMerge
        Range("K3:L3,K4:L4,K5:L5,K6:L6,K7:L7,K8:L8,K9:L9,K10:L10,K11:L11,M3:Q3,M4:Q4,M5:Q5,M6:Q6,M7:Q7,M8:Q8,M9:Q9,M10:Q10,M11:Q11,U3:V3,U4:V4,U5:V5,U6:V6,U7:V7,U8:V8,U9:V9,U10:V10,U11:V11").Select
        Range("K3:L3,K4:L4,K5:L5,K6:L6,K7:L7,K8:L8,K9:L9,K10:L10,K11:L11,M3:Q3,M4:Q4,M5:Q5,M6:Q6,M7:Q7,M8:Q8,M9:Q9,M10:Q10,M11:Q11,U3:V3,U4:V4,U5:V5,U6:V6,U7:V7,U8:V8,U9:V9,U10:V10,U11:V11").Merge
        
        'Default header Font reset
        Range("F3:W3").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 5855577
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .TintAndShade = 0
                .ThemeColor = xlThemeColorLight1
            End With
            With Selection.Font
                .Name = "Franklin Gothic Book"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            With Selection.Font
                .Name = "Franklin Gothic Book"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
    
        'Default data Font/Alignment Reset
        Range("F4:W11").Select
        With Selection.Font
            .Name = "Calibri"
            .Size = 11
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
        
EndMacro:
        Sheets("Traffic Workbook").Select
        Range("12:12").Select
        ActiveSheet.AutoFilterMode = False
        If Not ActiveSheet.AutoFilterMode Then
            ActiveSheet.Range("A12:AA12").AutoFilter
        End If
        Application.Goto Reference:=Range("A1"), Scroll:=True
        Range("A13").Select
        

    ActiveSheet.UsedRange
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
