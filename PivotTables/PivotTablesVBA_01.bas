Attribute VB_Name = "PivotTablesVBA_01"
' Sub PTs__CachesCountActiveWB()
' Sub PTs_CacheRecycle()
' Sub PTs_CountEachTab()
' Sub PTs_SelectPTIncFilters()
' Sub PTs_SelectPTExcFilters()

Sub PTs__CachesCountActiveWB()

    msg = "The number of pivot caches in " & Chr(34) & ActiveWorkbook.Name & Chr(34) & " is " & CStr(ActiveWorkbook.PivotCaches.Count) & "."
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' PTs__CachesCountActiveWB

Sub PTs_CacheRecycle()

    On Error GoTo TheEnd
    
    Dim p, WhichPT As Long
    Dim NameExists, Enabled As Boolean ' These defualt to False
    Dim ErrMsg, msg, PTs As String
    
    PTs = "Which pivot table's cache would you like to recycle? (Enter the number)" & vbCr
    
    ' Uncomment the following to enable
    Enabled = True
    
    If Enabled = False Then
    
        ErrMsg = "The procedure " & Chr(34) & "PTs_CacheRecycle" & Chr(34) & " is disabled."
    
    Else ' Enabled
    
        If ActiveSheet.PivotTables.Count = 0 Then
        
            ErrMsg = "No pivot tables on tab: " & Chr(34) & ActiveSheet.Name & Chr(34) & "."
        
        Else ' There are pivot tables on the active sheet
        
            If ActiveSheet.PivotTables.Count = 1 Then
            
                WhichPT = 1
                
            Else
                For p = 1 To ActiveSheet.PivotTables.Count
                
                    PTs = PTs & vbCr & CStr(p) & ": " & ActiveSheet.PivotTables(p).Name
                
                Next p
                
                WhichPT = InputBox(PTs, "Choose PT dialog.", 1)
            
            End If
            
            Range("A:D").Insert ' Make space for the pivot table
        
            For p = 1 To ActiveSheet.PivotTables.Count ' Recycling the p variable here ironically enough
            
                If ActiveSheet.PivotTables(p).Name = "Created from existing pivot" Then
                
                    NameExists = True
                    
                    Exit For
                
                End If
            
            Next p
            
            If NameExists Then
            
                ActiveSheet.PivotTables(WhichPT).PivotCache.CreatePivotTable _
                TableDestination:=Cells(1, 1)
            
            Else
            
                ActiveSheet.PivotTables(WhichPT).PivotCache.CreatePivotTable _
                TableDestination:=Cells(1, 1), _
                TableName:="Created from existing pivot"
            
            End If
            
        End If ' There are pivot tables on the active sheet
    
    End If
    
TheEnd:
    
    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number:" & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    Else
    
        If ErrMsg <> "" Then
        
            msg = ErrMsg
        
        Else
        
            msg = "Pivot table added!"
        
        End If
    
    End If
    
    Beep
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' PTs_CacheRecycle

Sub PTs_CountEachTab()

    Dim PTs As String
    Dim GrandTotal As Long
    
    PTs = "ActiveWorkbook.Name: " & vbCr & ActiveWorkbook.Name & vbCr
    
    For Each ws In ActiveWorkbook.Sheets
    
        PTs = PTs & vbCr & ws.Name & ": "
        PTs = PTs & CStr(ws.PivotTables.Count)
        
        GrandTotal = GrandTotal + ws.PivotTables.Count
    
    Next ws
    
    PTs = PTs & vbCr & "GRAND TOTAL: " & CStr(GrandTotal)
    
    Beep
    
    Debug.Print PTs
    
    MsgBox PTs

End Sub ' PTs_CountEachTab

Sub PTs_SelectPTIncFilters()

    ' Over-engineered for your comfort & convenience ;0)
    
    On Error GoTo TheEnd
    
    Dim p, WhichPT As Long
    Dim PTs, ErrMsg, msg As String
    PTs = "Enter the number of the pivot table (filters will be INCLUDED)." & vbCr
    
    If ActiveSheet.PivotTables.Count = 0 Then
    
        ErrMsg = "No pivot tables on tab: " & Chr(34) & ActiveSheet.Name & Chr(34) & "."
    
    Else
    
        If ActiveSheet.PivotTables.Count = 1 Then
        
            ActiveSheet.PivotTables(1).TableRange2.Select
        
        Else
        
            For p = 1 To ActiveSheet.PivotTables.Count
            
                PTs = PTs & vbCr & CStr(p) & ": " & ActiveSheet.PivotTables(p).Name
            
            Next p
            
            ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/inputbox-function
            ' InputBox(prompt, [ title ], [ default ], [ xpos ], [ ypos ], [ helpfile, context ])
            ' Only "prompt" is mandatory
            
           WhichPT = InputBox(PTs, "Choose PT dialog.", 1)
           
           ActiveSheet.PivotTables(WhichPT).TableRange2.Select
        
        End If
    
    End If

TheEnd:
    
    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number:" & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    Else
    
        If ErrMsg <> "" Then
        
            msg = ErrMsg
        
        Else
        
            ' msg = "Pivot"
        
        End If
    
    End If
    
    If msg <> "" Then
    
        Beep
        
        Debug.Print msg
        
        MsgBox msg
    
    End If

End Sub ' PTs_SelectPTIncFilters

Sub PTs_SelectPTExcFilters()

    ' Over-engineered for your comfort & convenience ;0)
    
    On Error GoTo TheEnd
    
    Dim p, WhichPT As Long
    Dim PTs, ErrMsg, msg As String
    PTs = "Enter the number of the pivot table (filters will be EXCLUDED)." & vbCr
    
    If ActiveSheet.PivotTables.Count = 0 Then
    
        ErrMsg = "No pivot tables on tab: " & Chr(34) & ActiveSheet.Name & Chr(34) & "."
    
    Else
    
        If ActiveSheet.PivotTables.Count = 1 Then
        
            ActiveSheet.PivotTables(1).TableRange1.Select
        
        Else
        
            For p = 1 To ActiveSheet.PivotTables.Count
            
                PTs = PTs & vbCr & CStr(p) & ": " & ActiveSheet.PivotTables(p).Name
            
            Next p
            
            ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/inputbox-function
            ' InputBox(prompt, [ title ], [ default ], [ xpos ], [ ypos ], [ helpfile, context ])
            ' Only prompt is mandatory
            
           WhichPT = InputBox(PTs, "Choose PT dialog.", 1)
           
           ActiveSheet.PivotTables(WhichPT).TableRange1.Select
        
        End If
    
    End If
    
TheEnd:
    
    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number:" & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    Else
    
        If ErrMsg <> "" Then
        
            msg = ErrMsg
        
        Else
        
            ' msg = "Pivot"
        
        End If
    
    End If
    
    If msg <> "" Then
    
        Beep
        
        Debug.Print msg
        
        MsgBox msg
    
    End If

End Sub ' PTs_SelectPTExcFilters

