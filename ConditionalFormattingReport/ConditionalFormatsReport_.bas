Attribute VB_Name = "ConditionalFormatsReport_"
' My ever useful ConditionalFormatsReport!
' Sub ConditionalFormatsCount()
' Sub ConditionalFormatsCountWithBreakdown()
' Sub ConditionalFormatsReport() -> *** MAIN EVENT ***
' Function WhichCols(MyRange As Range) As String
' Function GetRed(rgbCode As Long) As Long
' Function GetGreen(rgbCode As Long) As Long
' Function GetBlue(rgbCode As Long) As Long
' Private Sub SelectRange()
' Private Sub DeleteActiveSheet()
' Private Sub ObscureSettingValue()
' Private Sub ObscureSettingToggle()

Sub ConditionalFormatsCount()

    Dim GrandTotal As Long
    
    For Each ws In ActiveWorkbook.Sheets
    
        If ws.Type = 4 Then
        
            ' Chart
        
        Else
        
            GrandTotal = GrandTotal + ws.Cells.FormatConditions.Count
        
        End If
    
    Next ws
    
    MsgBox _
        "ActiveWorkbook.Name: " & Chr(34) & ActiveWorkbook.Name & Chr(34) & "." & _
        vbCr & vbCr & _
        "No of conditional formats: " & Format(GrandTotal, "#,##0") & "."

End Sub

Sub ConditionalFormatsCountWithBreakdown()

    Dim GrandTotal As Long
    Dim Details As String
    Details = "ActiveWorkbook.Name: " & Chr(34) & ActiveWorkbook.Name & Chr(34) & "." & vbCr & vbCr & "--- Breakdown by sheet ---"
    
    For Each ws In ActiveWorkbook.Sheets
    
        If ws.Type = 4 Then ' Chart
        
            Details = Details & vbCr & ws.Name & " (chart): n/a."
        
        Else ' Not a chart
        
            GrandTotal = GrandTotal + ws.Cells.FormatConditions.Count
        
            Details = Details & vbCr & ws.Name & ": " & Format(ws.Cells.FormatConditions.Count, "#,##0") & "."
        
        End If ' Not a chart
    
    Next ws
    
    Details = Details & vbCr & "GRAND TOTAL: " & Format(GrandTotal, "#,##0") & "."
    
    Beep
    
    Debug.Print Details

    MsgBox Details

End Sub ' ConditionalFormatsCountWithBreakdown

Sub ConditionalFormatsReport()

    ' There are 19 different properties of FormatCondition:
    ' https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition#properties
    
    ' Great setting!
    ' Application.AutoCorrect.AutoFillFormulasInLists = True
    
    On Error Resume Next
    
    StartTime = Now
    
    Application.Calculation = xlCalculationManual
    
    Dim r, c, ProcessedCFs As Long
    Dim ColumnHeaders, WriteToCell, ThisWB As String
    
    ColumnHeaders = ColumnHeaders & "Sheet"
    ColumnHeaders = ColumnHeaders & ", Applies to"
    ColumnHeaders = ColumnHeaders & ", Applies to (length)" ' Formula
    ColumnHeaders = ColumnHeaders & ", Applies to (columns)" ' Formula
    ColumnHeaders = ColumnHeaders & ", Type (value)"
    ColumnHeaders = ColumnHeaders & ", Type (desc)" ' Formula
    ColumnHeaders = ColumnHeaders & ", Operator (value)"
    ColumnHeaders = ColumnHeaders & ", Operator (desc)" ' Formula
    ColumnHeaders = ColumnHeaders & ", Formula1"
    ColumnHeaders = ColumnHeaders & ", Formula2"
    ColumnHeaders = ColumnHeaders & ", One"
    ColumnHeaders = ColumnHeaders & ", Stripe"
    ColumnHeaders = ColumnHeaders & ", Interior colour"
    
    If ThisWorkbook.Name <> ActiveWorkbook.Name Then
    
        ' We'll use this when ActiveWorkbook & ThisWorkbook are different see here ***
        ThisWB = "'" & ThisWorkbook.Name & "'!"
    
    End If
    
    Debug.Print WorksheetFunction.Rept("-", 40)
    Debug.Print "ActiveWorkbook.Name: " & Chr(34) & ActiveWorkbook.Name & Chr(34) & "."
    Debug.Print "ThisWorkbook.Name: " & Chr(34) & ThisWorkbook.Name & Chr(34) & "."
    Debug.Print "ThisWB: " & Chr(34) & ThisWB & Chr(34) & "."
    Debug.Print WorksheetFunction.Rept("-", 40)
    
    ' --- Separator ---
    
    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Sheets
    
        If ws.Name = "conditional_formats" Then
        
            ws.Delete: Exit For
        
        End If
    
    Next ws
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "conditional_formats"
    ActiveSheet.Tab.Color = RGB(0, 176, 240)
    
    [A2].Select
    ActiveWindow.FreezePanes = True
    
    Debug.Print "New conditional_formats tab added."
    
    For c = LBound(Split(ColumnHeaders, ", ")) To UBound(Split(ColumnHeaders, ", "))
    
        ' Writing headers to the first row here
        Cells(1, c + 1) = Split(ColumnHeaders, ", ")(c) ' c starts at zero
    
    Next c
    
    Debug.Print "Wrote headers to row one."
    
    ' Convert to list object
    ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=Cells(1, 1).CurrentRegion, _
        xlListObjectHasHeaders:=xlYes _
    ).Name = "ListOfConditionalFormats"
    
    Debug.Print "Converted to list object."
    
    ActiveSheet.ListObjects("ListOfConditionalFormats").TableStyle = "TableStyleLight14" ' Very minimal table style
    
    Cells.ColumnWidth = 10
    
    With Rows(1)
    
        .WrapText = True
        .VerticalAlignment = xlTop
    
    End With
    
    Application.DisplayAlerts = True
    
    Debug.Print "--- Output details of all conditional formats STARTS ---"
    
    r = 2
    
    For Each ws In ActiveWorkbook.Sheets
    
        If _
            ws.Name = "conditional_formats" Or _
            ws.Name = "also ignore this sheet" Or _
            ws.Name = "and this one" _
        Then
        
            ' Do nothing, not even increment the r variable
        
        Else ' It's a sheet we don't ignore
        
            If ws.Type = 4 Then ' Chart
            
                For c = LBound(Split(ColumnHeaders, ", ")) To UBound(Split(ColumnHeaders, ", "))
                
                    If Split(ColumnHeaders, ", ")(c) = "Sheet" Then
                    
                        Cells(r, c + 1) = ws.Name
                    
                    Else ' Split(ColumnHeaders, ", ")(c) <> "Sheet"
                    
                        Cells(r, c + 1) = "-"
                    
                    End If ' Split(ColumnHeaders, ", ")(c) <> "Sheet"
                
                Next c
                
                r = r + 1
            
            Else ' Not a Chart
            
                If ws.Cells.FormatConditions.Count = 0 Then
                
                    For c = LBound(Split(ColumnHeaders, ", ")) To UBound(Split(ColumnHeaders, ", "))
                    
                        If Split(ColumnHeaders, ", ")(c) = "Sheet" Then
                        
                            Cells(r, c + 1) = ws.Name
                        
                        Else ' Split(ColumnHeaders, ", ")(c) <> "Sheet"
                        
                            Cells(r, c + 1) = "-"
                        
                        End If ' Split(ColumnHeaders, ", ")(c) <> "Sheet"
                    
                    Next c
                    
                    r = r + 1
                
                Else ' ws.Cells.FormatConditions.Count > 0
                
                    For i = 1 To ws.Cells.FormatConditions.Count
                    
                        ' Debug.Print "Sheet name: " & Chr(34) & ws.Name & Chr(34) & ". Format condition number: " & CStr(i) & "."
                        
                        Set CondForm = ws.Cells.FormatConditions.Item(i)
                        
                        For c = LBound(Split(ColumnHeaders, ", ")) To UBound(Split(ColumnHeaders, ", "))
                        
                           Select Case Split(ColumnHeaders, ", ")(c)
                           
                                Case "Sheet"
                                
                                    Cells(r, c + 1) = ws.Name
                           
                                Case "Applies to"
                                
                                    Cells(r, c + 1) = Replace(CondForm.AppliesTo.Address, "$", "")
                                
                                Case "Applies to (length)"
                                
                                    ' Do nothing: We'll do this as a formula
                                
                                Case "Applies to (columns)"
                                
                                    ' Do nothing: We'll do this as a formula
                                
                                Case "Type (value)"
                                
                                    Cells(r, c + 1) = CondForm.Type
                                
                                Case "Type (desc)"
                                
                                    ' 14 of these
                                    ' https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditiontype
                                    ' Do nothing: We'll do this as a formula
                                
                                Case "Operator (value)"
                                
                                    WriteToCell = CondForm.Operator
                                    
                                    If Err.Number > 0 Then
                                    
                                        Err.Clear
                                        
                                        WriteToCell = "n/a"
                                        
                                    End If
                                    
                                    Cells(r, c + 1) = WriteToCell
                                
                                Case "Operator (desc)"
                                
                                    ' 8 of these:
                                    ' https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditionoperator
                                    ' Do nothing: We'll do this as a formula
                                
                                Case "Formula1"
                                
                                    WriteToCell = Chr(39) & CondForm.Formula1 ' Notice the apostrophe
                                    
                                    If Err.Number > 0 Then
                                    
                                        Err.Clear
                                        
                                        WriteToCell = "n/a"
                                    
                                    End If
                                    
                                    Cells(r, c + 1) = WriteToCell
                                    
                                Case "Formula2"
                                    
                                    WriteToCell = Chr(39) & CondForm.Formula2 ' Notice the apostrophe
                                    
                                    If Err.Number > 0 Then
                                    
                                        Err.Clear
                                        
                                        WriteToCell = "n/a"
                                    
                                    End If
                                    
                                    Cells(r, c + 1) = WriteToCell
                                    
                                Case "Interior colour"
                                
                                    If CondForm.Interior.ColorIndex = -4142 Then
                                    
                                        WriteToCell = "No fill"
                                    
                                    Else
                                    
                                        WriteToCell = "RGB(" & _
                                        GetRed(CondForm.Interior.Color) & ", " & _
                                        GetGreen(CondForm.Interior.Color) & ", " & _
                                        GetBlue(CondForm.Interior.Color) & _
                                        ")"
                                    
                                    End If
                                    
                                    If Err.Number > 0 Then
                                        
                                        Err.Clear
                                        WriteToCell = "n/a"
                                    
                                    End If
                                    
                                    Cells(r, c + 1) = WriteToCell
                                    
                                    If WriteToCell <> "n/a" And WriteToCell <> "No fill" Then
                                    
                                        Cells(r, c + 1).Interior.Color = CondForm.Interior.Color
                                        
                                    End If
                                    
                                Case Else
                                
                                    ' Do nothing
                           
                           End Select
                        
                        Next c
                        
                        Set CondForm = Nothing
                        
                        r = r + 1
                    
                    Next i
                
                End If ' ws.Cells.FormatConditions.Count > 0
            
            End If ' Not a Chart
        
        End If ' It's a sheet we don't ignore
    
    Next ws

    Debug.Print "--- Output details of all conditional formats ENDS ---"
    
    Debug.Print "--- Formulas START ---"
    
    ' Simple formulas
    
    Range("ListOfConditionalFormats[One]") = "=1"
    Debug.Print "Populated column: " & Chr(34) & "One" & Chr(34) & "."
    
    Range("ListOfConditionalFormats[Applies to (length)]") = _
    "=IF([@[Applies to]]=" & Chr(34) & "-" & Chr(34) & ", " & Chr(34) & "-" & Chr(34) & ", LEN([@[Applies to]]))"
    Debug.Print "Populated column: " & Chr(34) & "Applies to (length)" & Chr(34) & "."
    
    ' If you're running this procedure from another file, you need to reference ThisWorkbook in the following formula
    Range("ListOfConditionalFormats[Applies to (columns)]") = "=" & ThisWB & "WhichCols([@[Applies to]])" ' ***
    Debug.Print "Populated column: " & Chr(34) & "Applies to (columns)" & Chr(34) & "."
    
    ' More complicated formulas
    
    theFormula = "=IF("
    theFormula = theFormula & "OFFSET([@Stripe],-1,0)=" & Chr(34) & "Stripe" & Chr(34)
    theFormula = theFormula & ", 0"
    theFormula = theFormula & ", IF("
    theFormula = theFormula & "OFFSET([@Sheet],-1,0)=[@Sheet]"
    theFormula = theFormula & ", OFFSET([@Stripe],-1,0)"
    theFormula = theFormula & ", IF("
    theFormula = theFormula & "OFFSET([@Stripe],-1,0)=1, 0, 1)))"
    
    Range("ListOfConditionalFormats[Stripe]") = theFormula
    Debug.Print "Populated column: " & Chr(34) & "Stripe" & Chr(34) & "."
    
    theFormula = "=IF("
    theFormula = theFormula & "[@[Type (value)]] = " & Chr(34) & "-" & Chr(34)
    theFormula = theFormula & ", " & Chr(34) & "-" & Chr(34)
    theFormula = theFormula & ", INDEX({"
    theFormula = theFormula & Chr(34) & "xlCellValue" & Chr(34) & ", " ' 01 of 14: 1
    theFormula = theFormula & Chr(34) & "xlExpression" & Chr(34) & ", " ' 02 of 14: 2
    theFormula = theFormula & Chr(34) & "xlColorScale" & Chr(34) & ", " ' 03 of 14: 3
    theFormula = theFormula & Chr(34) & "xlDataBar" & Chr(34) & ", " ' 04 of 14: 4
    theFormula = theFormula & Chr(34) & "xlTop10" & Chr(34) & ", " ' 05 of 14: 5
    theFormula = theFormula & Chr(34) & "xlIconSet" & Chr(34) & ", " ' 06 of 14: 6
    theFormula = theFormula & Chr(34) & "xlUniqueValues" & Chr(34) & ", " ' 07 of 14: 8
    theFormula = theFormula & Chr(34) & "xlTextString" & Chr(34) & ", " ' 08 of 14: 9
    theFormula = theFormula & Chr(34) & "xlBlanksCondition" & Chr(34) & ", " ' 09 of 14: 10
    theFormula = theFormula & Chr(34) & "xlTimePeriod" & Chr(34) & ", " ' 10 of 14: 11
    theFormula = theFormula & Chr(34) & "xlAboveAverageCondition" & Chr(34) & ", " ' 11 of 14: 12
    theFormula = theFormula & Chr(34) & "xlNoBlanksCondition" & Chr(34) & ", " ' 12 of 14: 13
    theFormula = theFormula & Chr(34) & "xlErrorsCondition" & Chr(34) & ", " ' 13 of 14: 16
    theFormula = theFormula & Chr(34) & "xlNoErrorsCondition" & Chr(34) ' 14 of 14: 17
    theFormula = theFormula & "}, MATCH([@[Type (value)]], {1,2,3,4,5,6,8,9,10,11,12,13,16,17}, 0))" ' 7, 14 & 15 missing
    theFormula = theFormula & ")"
    
    Range("ListOfConditionalFormats[Type (desc)]") = theFormula
    Debug.Print "Populated column: " & Chr(34) & "Type (desc)" & Chr(34) & "."
    
    theFormula = "=IFERROR(IF("
    theFormula = theFormula & "[@[Operator (value)]] = " & Chr(34) & "-" & Chr(34)
    theFormula = theFormula & ", " & Chr(34) & "-" & Chr(34)
    theFormula = theFormula & ", INDEX({"
    theFormula = theFormula & Chr(34) & "xlBetween" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlNotBetween" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlEqual" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlNotEqual" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlGreater" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlLess" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlGreaterEqual" & Chr(34) & ", "
    theFormula = theFormula & Chr(34) & "xlLessEqual" & Chr(34)
    theFormula = theFormula & "}, MATCH([@[Operator (value)]], {1,2,3,4,5,6,7,8}, 0)"
    theFormula = theFormula & ")), " & Chr(34) & "n/a" & Chr(34) & ")"
    
    Range("ListOfConditionalFormats[Operator (desc)]") = theFormula
    Debug.Print "Populated column: " & Chr(34) & "Operator (desc)" & Chr(34) & "."
    
    ' Destroy formulas
    ActiveSheet.ListObjects(1).DataBodyRange.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False ' Clear marching ants
    
    Debug.Print "Formulas destroyed."
    
    Debug.Print "--- Formulas END ---"
    
    Debug.Print "--- Search & replace in 3 columns STARTS ---"
    
    For Each Item In Split("Operator (value), Formula1, Formula2", ", ")
    
        Range("ListOfConditionalFormats[" & Item & "]").Select
        
        If WorksheetFunction.CountBlank(Selection) > 0 Then
        
            Selection.Replace What:="", Replacement:="-"
        
        End If
        
        Debug.Print "Blanks (if any) were replaced in column: " & Chr(34) & Item & Chr(34) & "."
    
    Next Item
    
    Debug.Print "--- Search & replace in 3 columns ENDS ---"
    
    Debug.Print "--- Conditional formats START ---"
    
    Range("ListOfConditionalFormats[[Sheet]:[Stripe]]").Select ' Not the last col which is "Interior colour"
    
    theFormula = "=" & Replace(Cells(2, WorksheetFunction.Match("stripe", Rows(1), 0)).Address, "$2", "2") & " = 0"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=theFormula
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.Color = RGB(255, 255, 204) ' Light yellow
    Selection.FormatConditions(1).StopIfTrue = False
    Debug.Print "Light yellow stripes added."
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=Replace(theFormula, " = 0", " = 1")
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.Color = RGB(204, 236, 255) ' Light blue
    Selection.FormatConditions(1).StopIfTrue = False
    Debug.Print "Light blue stripes added."
    
    Range("ListOfConditionalFormats[Applies to (length)]").Select
    
    Selection.FormatConditions.Add Type:=xlExpression, _
    Formula1:="=AND(" & _
    Replace(Selection.Cells(1).Address, "$", "") & "<>" & Chr(34) & "-" & Chr(34) & ", " & _
    Replace(Selection.Cells(1).Address, "$", "") & ">25" & _
    ")"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.Color = RGB(255, 255, 0) ' Yellow
    Selection.FormatConditions(1).StopIfTrue = False
    Debug.Print "Any conditional format where the length of " & Chr(34) & "Applies to" & Chr(34) & " > 25 is now coloured yellow."
    ' Applies to will be something like A1:E1000. What we're looking for is something like this: A1:E5, A7:E8, A10:E15, etc
    
    Range("ListOfConditionalFormats[Interior colour]").Select
    
    Selection.FormatConditions.Add Type:=xlExpression, _
    Formula1:="=OR(" & _
    Replace(Selection.Cells(1).Address, "$", "") & "=" & Chr(34) & "-" & Chr(34) & ", " & _
    Replace(Selection.Cells(1).Address, "$", "") & "=" & Chr(34) & "n/a" & Chr(34) & ", " & _
    Replace(Selection.Cells(1).Address, "$", "") & "=" & Chr(34) & "no fill" & Chr(34) & _
    ")"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.ColorIndex = -4142
    Selection.FormatConditions(1).StopIfTrue = False
    Debug.Print "No fills & hyphens' ColorIndex set to -4142 (no fill)."
    
    Debug.Print "--- Conditional formats END ---"
    
    Debug.Print "--- Autofit certain columns then delete a couple STARTS ---"
    
    AutofitThese = "" ' So subsequent items line up nicely
    AutofitThese = AutofitThese & "sheet, "
    AutofitThese = AutofitThese & "applies to (columns), "
    AutofitThese = AutofitThese & "operator (desc), "
    AutofitThese = AutofitThese & "type (desc), "
    AutofitThese = AutofitThese & "interior colour, "
    AutofitThese = AutofitThese & "one, "
    AutofitThese = AutofitThese & "stripe"
    
    For Each Item In Split(AutofitThese, ", ")
    
        Columns(WorksheetFunction.Match(Item, Rows(1), 0)).EntireColumn.AutoFit
        Debug.Print "Autofit column: " & Chr(34) & Item & Chr(34) & "."
    
    Next Item
    
    ' Make "Formula1" column wider
    
    Columns(WorksheetFunction.Match("formula1", Rows(1), 0)).ColumnWidth = _
    5 * Columns(WorksheetFunction.Match("formula1", Rows(1), 0)).ColumnWidth
    
    Debug.Print "Made " & Chr(34) & "Formula1" & Chr(34) & " column 5x wider than it was."
    
    ' Deletions
    
    For Each Item In Split("type (value), operator (value)", ", ")
    
        ' We have the descs for these so we don't need the numbers
        Columns(WorksheetFunction.Match(Item, Rows(1), 0)).Delete
        Debug.Print "Deleted column: " & Chr(34) & Item & Chr(34) & "."
    
    Next Item
    
    Debug.Print "--- Autofit certain columns then delete a couple ENDS ---"
    
    Debug.Print "--- Last few bits STARTS ---"
    
    Range("1:3").Insert
    
    Debug.Print "3 rows were inserted at the top."
    
    With [A1]
    
        .Value = "ActiveWB: " & ActiveWorkbook.Name
        
        With .Font
        
            .Bold = True
            .Color = RGB(255, 0, 0)
        
        End With
    
    End With
    
    Debug.Print "ActiveWorkbook.Name written to [A1]."
    
    ' --- Separator ---
    
    With [A2]
    
        .Value = "ThisWB: " & ThisWorkbook.Name
        
        With .Font
        
            .Bold = True
            .Color = RGB(255, 0, 0)
        
        End With
    
    End With
    
    Debug.Print "ThisWorkbook.Name written to [A2]."
    
    ' No of cells in the referenced columns that aren't blank or "-"
    
    theFormula = "=COUNTIF(INDIRECT("
    theFormula = theFormula & Chr(34) & "ListOfConditionalFormats[" & Chr(34) & " & A4 & " & Chr(34) & "]" & Chr(34) & "), "
    theFormula = theFormula & Chr(34) & "<>" & Chr(34)
    theFormula = theFormula & ") - COUNTIF(INDIRECT("
    theFormula = theFormula & Chr(34) & "ListOfConditionalFormats[" & Chr(34) & " & A4 & " & Chr(34) & "]" & Chr(34) & "), "
    theFormula = theFormula & Chr(34) & "-" & Chr(34)
    theFormula = theFormula & ")"
    
    Range( _
        Cells(3, 1), _
        Cells(3, ActiveSheet.ListObjects(1).ListColumns.Count) _
    ).Select
    
    With Selection
    
        .Interior.Color = RGB(255, 255, 0)
        .Formula = theFormula
        .NumberFormat = "#,##0"
    
    End With
    
    Debug.Print "Added formulas on row 3."
    
    ' Page setup
    
    Application.PrintCommunication = False
    
    With ActiveSheet.PageSetup
        
        .Orientation = xlLandscape
        .CenterHeader = "Page &P of &N"
        .PrintTitleRows = "$1:$4"
    
    End With
    
    Application.PrintCommunication = True
    
    Debug.Print "Page setup done."
    
    Debug.Print "--- Last few bits ENDS ---"
    
    Debug.Print "--- Add buttons STARTS ---"

    ' First button
    
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
    
        [D1] = "Didn't"
        
        Debug.Print "Didn't add Delete button 'cos it would've been externally linked."
    
    Else ' ActiveWorkbook = ThisWorkbook
    
        AddDeleteButton = 1
        
        If AddDeleteButton = 1 Then
        
            ButtonText = "Delete"
            
            Set theRange = Range("D1:D2")
            
            ActiveSheet.Buttons.Add(theRange.Left, theRange.Top, theRange.Width, theRange.Height).Select
            Selection.OnAction = "DeleteActiveSheet"
            Selection.Caption = ButtonText
            
            With Selection.Characters(Start:=1, Length:=Len(ButtonText)).Font
            
                .FontStyle = "Bold"
            
            End With
            
            Debug.Print "Added Delete button."
        
        End If
    
    End If
    
    ' Second button
    
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
    
        [E1] = "Add"
        
        Debug.Print "Didn't add Refresh button 'cos it would've been externally linked."
    
    Else ' ActiveWorkbook.Name = ThisWorkbook.Name
    
        ButtonText = "Refresh"
        
        Set theRange = Range("E1:E2")
        
        ActiveSheet.Buttons.Add(theRange.Left, theRange.Top, theRange.Width, theRange.Height).Select
        Selection.OnAction = "ConditionalFormatsReport"
        Selection.Caption = ButtonText
        
        With Selection.Characters(Start:=1, Length:=Len(ButtonText)).Font
        
            .FontStyle = "Bold"
        
        End With
        
        Debug.Print "Added Refresh button."
    
    End If ' ActiveWorkbook.Name = ThisWorkbook.Name
    
    ' Third button
    
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then
    
        [F1] = "Buttons"
        
        Debug.Print "Didn't add Select button 'cos it would've been externally linked."
    
    Else ' ActiveWorkbook.Name = ThisWorkbook.Name
    
        ButtonText = "Select"
        
        Set theRange = Range("F1:F2")
        
        ActiveSheet.Buttons.Add(theRange.Left, theRange.Top, theRange.Width, theRange.Height).Select
        Selection.OnAction = "SelectRange" ' Private
        Selection.Caption = ButtonText
        
        With Selection.Characters(Start:=1, Length:=Len(ButtonText)).Font
        
            .FontStyle = "Bold"
        
        End With
        
        Debug.Print "Added Select button."
    
    End If ' ActiveWorkbook.Name = ThisWorkbook.Name
    
    Debug.Print "--- Add buttons ENDS ---"

TheEnd:
    
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number: " & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    Else
    
        msg = "Done! Time taken: " & Format(Now - StartTime, "hh:mm:ss") & "."
    
    End If
    
    Range([A1], ActiveCell.SpecialCells(xlLastCell)).Select
    
    ThereAreFormulas = False
    
    For Each theCell In Selection.Cells
    
        If theCell.HasFormula Then
        
            ThereAreFormulas = True: Exit For
        
        End If
    
    Next theCell
    
    If ThereAreFormulas Then
    
        Selection.SpecialCells(xlCellTypeFormulas, 23).Select ' Select formulas
    
    End If
    
    Beep
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' ConditionalFormatsReport

Function WhichCols(MyRange As Range) As String
    
    WhichCols = MyRange.Value ' eg $A$1:$A$1000, $B$1:$E$1000

    OutputThis = ""
    
    RemoveTheseChars = " 0123456789$"
    
    For i = 1 To Len(RemoveTheseChars)
    
        WhichCols = Replace(WhichCols, Mid(RemoveTheseChars, i, 1), "")
    
    Next i
    
    ' You'll end up with something like A:A,B:E
    
    For i = LBound(Split(WhichCols, ",")) To UBound(Split(WhichCols, ","))

        If InStr(Split(WhichCols, ",")(i), ":") > 0 Then
        
            If Split(Split(WhichCols, ",")(i), ":")(0) = Split(Split(WhichCols, ",")(i), ":")(1) Then
            
                If OutputThis = "" Then
                
                    OutputThis = Split(Split(WhichCols, ",")(i), ":")(0)
                
                Else
                
                    OutputThis = OutputThis & ", " & Split(Split(WhichCols, ",")(i), ":")(0)
                
                End If
            
            Else ' Split(Split(WhichCols, ",")(i), ":")(0) <> Split(Split(WhichCols, ",")(i), ":")(1)
            
                If OutputThis = "" Then
                
                    OutputThis = Split(WhichCols, ",")(i)
                
                Else
                
                    OutputThis = OutputThis & ", " & Split(WhichCols, ",")(i)
                
                End If
            
            End If
        
        Else ' No colon
        
            If OutputThis = "" Then
            
                OutputThis = Split(WhichCols, ",")(i)
            
            Else
            
                OutputThis = OutputThis & Split(WhichCols, ",")(i)
            
            End If
        
        End If
    
    Next i
    
    WhichCols = OutputThis

End Function ' WhichCols

Function GetRed(rgbCode As Long) As Long
    
    GetRed = rgbCode Mod 256

End Function

Function GetGreen(rgbCode As Long) As Long
    
    GetGreen = (rgbCode \ 256) Mod 256

End Function

Function GetBlue(rgbCode As Long) As Long
    
    GetBlue = rgbCode \ 65536

End Function

Private Sub SelectRange()

    On Error GoTo TheEnd
    
    ErrMsg = ""
    
    ' --- Error traps START ---
    
    If ActiveSheet.Name <> "conditional_formats" Then
    
        ErrMsg = "You need to be on the 'conditional_formats' tab to run the 'SelectRange' macro."
        ErrMsg = Replace(ErrMsg, Chr(39), Chr(34))
    
    End If
    
    If ErrMsg = "" Then
    
        If ActiveSheet.ListObjects.Count = 0 Then
        
            ErrMsg = "ActiveSheet.ListObjects.Count = 0 :0("
        
        End If
    
    End If
    
    If ErrMsg = "" Then
    
        If _
            WorksheetFunction.CountIf(ActiveSheet.ListObjects(1).HeaderRowRange, "sheet") = 0 _
            Or _
            WorksheetFunction.CountIf(ActiveSheet.ListObjects(1).HeaderRowRange, "applies to") = 0 _
        Then
        
            ErrMsg = "" ' So subsequent items line up nicely
            ErrMsg = ErrMsg & "At least one of the following 2 columns "
            ErrMsg = ErrMsg & "is missing from list object " & Chr(34) & ActiveSheet.ListObjects(1).Name & Chr(34) & ": "
            ErrMsg = ErrMsg & Chr(34) & "Sheet" & Chr(34) & ", " & Chr(34) & "Applies to" & Chr(34) & "."
        
        End If
    
    End If
    
    If ErrMsg = "" Then
    
        If Application.Intersect( _
            Selection.Cells(1), _
            Range("ListOfConditionalFormats") _
        ) Is Nothing Then
        
            ErrMsg = "" ' So subsequent items line up nicely
            ErrMsg = ErrMsg & "The first cell in your selection ["
            ErrMsg = ErrMsg & Replace(Selection.Cells(1).Address, "$", "")
            ErrMsg = ErrMsg & "] DOES NOT intersect with 'ListOfConditionalFormats'."
            ErrMsg = Replace(ErrMsg, Chr(39), Chr(34))
        
        End If
    
    End If
    
    ' --- Error traps END ---
    
    ' --- Actual selecting STARTS here ---
    
    If ErrMsg = "" Then
    
        ' You can start selecting
        
        SheetSelect = Cells( _
            Selection.Cells(1).Row, _
            WorksheetFunction.Match( _
                "sheet", _
                ActiveSheet.ListObjects(1).HeaderRowRange, _
                0 _
            ) _
        )
        Debug.Print Replace("SheetSelect: '" & SheetSelect & "'.", Chr(39), Chr(34))
        
        ' #Feb2025
        ' For some season calling the following "SelectRange" resulted in an error
        ' 'Cos "SelectRange" is the name of this procedure!
        
        RangeSelect = Cells( _
            Selection.Cells(1).Row, _
            WorksheetFunction.Match( _
                "applies to", _
                ActiveSheet.ListObjects(1).HeaderRowRange, _
                0 _
            ) _
        )
        Debug.Print Replace("RangeSelect: '" & RangeSelect & "'.", Chr(39), Chr(34))
        
        Sheets(SheetSelect).Select
        
        Range(RangeSelect).Select
    
    End If
    
    ' --- Actual selecting ENDS here ---

TheEnd:

    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number:" & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    Else
    
        If ErrMsg <> "" Then
        
            msg = ErrMsg
        
        Else
        
            msg = "Range was selected"
        
        End If
    
    End If
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' SelectRange

Private Sub DeleteActiveSheet()
    
    Dim msg As String
    Dim answer As Byte
    
    answer = MsgBox( _
        "Delete tab " & Chr(34) & ActiveSheet.Name & Chr(34) & " in file " & Chr(34) & ActiveWorkbook.Name & Chr(34) & "? " & _
        "Default is " & Chr(34) & "no" & Chr(34) & ".", _
        vbYesNo + vbDefaultButton2 _
    )
    Select Case answer
    
        Case 6
        
            Application.DisplayAlerts = False
            ActiveSheet.Delete
            Application.DisplayAlerts = True
            
            msg = "Tab deleted."
    
        Case 7
        
            msg = "You clicked " & Chr(34) & "no" & Chr(34) & "."
    
        Case Else
        
            ' You can't cancel
    
    End Select
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' DeleteActiveSheet

Private Sub ObscureSettingValue()

    ' Via the menus: File > Options > Proofing > Click the AutoCorrect Options... button ...
    ' ... Tab: AutoFormat As You Type > Tick box: Fill formulae in tables to create calculated columns
    
    MsgBox "Application.AutoCorrect.AutoFillFormulasInLists: " & CStr(Application.AutoCorrect.AutoFillFormulasInLists) & "."

End Sub

Private Sub ObscureSettingToggle()

    Application.AutoCorrect.AutoFillFormulasInLists = Not Application.AutoCorrect.AutoFillFormulasInLists
    
    MsgBox "Application.AutoCorrect.AutoFillFormulasInLists: " & CStr(Application.AutoCorrect.AutoFillFormulasInLists) & "."

End Sub
