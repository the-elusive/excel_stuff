Attribute VB_Name = "ArraysTestBed"
Private Sub CopyListObjToMemory()
    
    ' On Error GoTo TheEnd
    
    ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
    ' IMPORTANT: The VarType function never returns the value for vbArray by itself.
    ' It's always added to some other value to indicate an array of a particular type.
    ' For example, the value returned for an array of integers is calculated as vbInteger (2) + vbArray (8192), or 8194.
    ' vbArray = 8192
    ' vbVariant = 12 > Used only with arrays of variants
    ' 8192 + 12 = 8204
    ' I was going to use the CallByName function in the appropriate places but decided against it
    ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/callbyname-function
    
    Debug.Print "--- CopyListObjToMemory STARTS ---"
    
    Dim StartTime As Date
    Dim CopyPasteTook As Date
    Dim myArray As Variant
    Dim answer As Byte
    
    answer = MsgBox( _
        prompt:=Replace( _
            "Copy list object to 'Destination':" & vbCr & _
            "Yes: All at once" & vbCr & _
            "No: One by one", _
            Chr(39), _
            Chr(34) _
        ), _
        Title:="Dialog 1 of 1", _
        Buttons:=vbYesNo + vbDefaultButton1 _
    )
    StartTime = Now
    Sheets("Source").Select
    
    ' Did consider using CallByName here
    myArray = ActiveSheet.Range("SourceListObj[#All]").Value
    
    ' Dimensions: 1 = Rows, 2 = Columns
    Dim RowsLower As Long: RowsLower = LBound(myArray, 1)
    Dim RowsUpper As Long: RowsUpper = UBound(myArray, 1)
    Dim ColsLower As Long: ColsLower = LBound(myArray, 2)
    Dim ColsUpper As Long: ColsUpper = UBound(myArray, 2)
    Dim r As Long, c As Long
    
    ' --- #These will be included in the final dialog ---
    Dim msg As String
    msg = msg & "  VarType(myArray) = " & Format(VarType(myArray), "#,##0") & "." & vbCr
    msg = msg & "  LBound(myArray, 1) ie RowsLower = " & Format(RowsLower, "#,##0") & "." & vbCr
    msg = msg & "  UBound(myArray, 1) ie RowsUpper = " & Format(RowsUpper, "#,##0") & "." & vbCr
    msg = msg & "  LBound(myArray, 2) ie ColsLower = " & Format(ColsLower, "#,##0") & "." & vbCr
    msg = msg & "  UBound(myArray, 2) ie ColsUpper = " & Format(ColsUpper, "#,##0") & "."
    
    Sheets("Destination").Select
    Cells.Clear
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ' Also considered using CallByName here
    Select Case answer
    
        Case 6 ' Yes: All at once (Fast)
            Cells(1, 1).Resize(RowsUpper, ColsUpper).Value = myArray
    
        Case 7 ' No: One by one (Slow)
            
            For r = RowsLower To RowsUpper
                For c = ColsLower To ColsUpper
                    Cells(r, c).Value = myArray(r, c)
                Next c
            Next r
            
    End Select
    
    CopyPasteTook = Now - StartTime
    
    Sheets("Source").Select
    
TheEnd:

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number > 0 Then
    
        msg = _
            "--- Err ---" & vbCr & _
            ".Number : " & CStr(Err.Number) & vbCr & _
            ".Description : " & Err.Description
        
        Err.Clear
    
    Else

        msg = _
            "  Done!" & vbCr & _
            "  answer (6 = All at once, 7 = one by one): " & CStr(answer) & "." & vbCr & _
            "  CopyPasteTook: " & Format(CopyPasteTook, "hh:mm:ss") & "." & vbCr & _
            "  Total time taken: " & Format(Now - StartTime, "hh:mm:ss") & "." & vbCr & vbCr & _
            msg ' See #These
    
    End If
    
    Debug.Print msg
    Debug.Print "--- CopyListObjToMemory ENDS ---" & vbCr
    MsgBox msg

End Sub ' CopyListObjToMemory
