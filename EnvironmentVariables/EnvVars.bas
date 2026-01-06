Attribute VB_Name = "EnvVars"
' Sub DummyMacro()
' Sub DoReset()
' Sub DoListEnvironmentVariables()
' Sub DoPathVariableOnly()
' Sub GetThisWorkbookPath()
' Sub LineFeedChars()
' Sub AddColToStartOfLO()

Sub DummyMacro()

    MsgBox "This macro does nothing."

End Sub

Sub DoReset()

    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Sheets
    
        If ws.Name <> "Home" Then
        
            ws.Delete
        
        End If
    
    Next ws
    
    Application.DisplayAlerts = True

End Sub ' DoReset

Sub DoListEnvironmentVariables()
    
    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Sheets
    
        If ws.Name = "EnvironmentVariables" Then
        
            ws.Delete
        
        End If
    
    Next
    
    Sheets.Add.Name = "EnvironmentVariables": Cells.Interior.Color = RGB(72, 61, 139)
    
    [A1] = "Var"
    
    Application.DisplayAlerts = True
    
    Dim strEnviron As String
    
    Dim i As Long
    
    For i = 1 To 255
        
        strEnviron = Environ(i)
        
        If LenB(strEnviron) = 0& Then
        
            Exit For
            
        End If
        
        If InStr(strEnviron, ";") > 0 Then
        
            Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1) = Replace(Replace(strEnviron, ";", ";" & vbLf), "=", "=" & vbLf)
            
        Else
        
            Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1) = strEnviron
        
        End If
    
    Next
    
    ' Convert to list object
    
    With ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=Cells(1, 1).CurrentRegion, _
        XlListObjectHasHeaders:=xlYes)
        
        .Name = "ListObj_EnvVars"
        .TableStyle = "TableStyleMedium7"
        .Range.Interior.ColorIndex = -4142
        
    End With
    
    Columns.AutoFit
    Rows.AutoFit
    
    Columns(1).Insert
    Columns(1).ColumnWidth = 2
    Rows(1).Insert
    
    Range("A3").Select
    ActiveWindow.FreezePanes = True
    
    ' Bold everything up to & including the first equals sign (if it exists)
    
    ActiveSheet.ListObjects(1).DataBodyRange.Select
    
    For Each theCell In Selection.Cells
    
        If InStr(theCell, "=") > 0 Then
        
            theCell.Characters(1, InStr(theCell, "=")).Font.Bold = True
        
        End If
    
    Next theCell
    
    ' Awesome: https://www.microsoft.com/en-us/download/details.aspx?id=50745
    Application.CommandBars.ExecuteMso "Filter"
    
    MsgBox "Done."

End Sub ' DoListEnvironmentVariables

Sub DoPathVariableOnly()

    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Sheets
    
        If ws.Name = "PathVariableOnly" Then
        
            ws.Delete
        
        End If
    
    Next
    
    Sheets.Add.Name = "PathVariableOnly": Cells.Interior.Color = RGB(72, 61, 139)
    
    [A1] = "Item"
    
    Application.DisplayAlerts = True

    Dim strEnviron As String
    
    Dim i As Long
    
    For i = 1 To 255
        
        strEnviron = Environ(i)
        
        If LenB(strEnviron) = 0& Then
        
            Exit For
            
        End If
        
        If Left(LCase(strEnviron), 5) = "path=" Then
        
            SplitThis = Right(strEnviron, Len(strEnviron) - 5)
            
            For Each Item In Split(SplitThis, ";")
            
                Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1) = Item
            
            Next
        
        End If
    
    Next
    
    ' Convert to list object
    
    With ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=Cells(1, 1).CurrentRegion, _
        XlListObjectHasHeaders:=xlYes)
        
        .Name = "ListObj_PathOnly"
        .TableStyle = "TableStyleMedium7"
        .Range.Interior.ColorIndex = -4142
        
    End With
    
    Columns(1).EntireColumn.AutoFit
    Columns(1).Insert
    Columns(1).ColumnWidth = 2
    
    Rows(1).Insert
    
    ' Awesome: https://www.microsoft.com/en-us/download/details.aspx?id=50745
    ActiveSheet.ListObjects(1).DataBodyRange.Select
    Application.CommandBars.ExecuteMso "Filter"
    
    MsgBox "Done."

End Sub ' DoPathVariableOnly

Sub GetThisWorkbookPath()

    MsgBox "ThisWorkbook.Path:" & vbCr & ThisWorkbook.Path

End Sub ' GetThisWorkbookPath

Sub LineFeedChars()

    ' Which delimiter was cause separate values to be displayed one atop the other in a cell?
    Sheets.Add
    
    [A1] = "vbCrLf" & vbCrLf & "vbCrLf" & vbCrLf & "vbCrLf" ' Good
    [A2] = "vbCr" & vbCr & "vbCr" & vbCr & "vbCr"
    [A3] = "vbLf" & vbLf & "vbLf" & vbLf & "vbLf" ' Also good

End Sub ' LineFeedChars

Sub AddColToStartOfLO()

    If ActiveSheet.ListObjects.Count = 0 Then
    
        MsgBox "ActiveSheet.ListObjects.Count = 0 :0("
    
    Else
    
        ActiveSheet.ListObjects(1).ListColumns.Add 1
    
    End If

End Sub ' AddColToStartOfLO
