Attribute VB_Name = "ListObjs_RefreshMethods"
' https://community.fabric.microsoft.com/t5/Power-Query/Querytable-refresh-vs-ListObject-Refresh/td-p/2890348
' Sub ListObjs_RefreshFirstMethod1() -> ActiveSheet.ListObjects(1).Refresh
' Sub ListObjs_RefreshFirstMethod2() -> ActiveSheet.ListObjects(1).QueryTable.Refresh
' Sub ListObjs_RefreshFirstMethod3() -> ActiveWorkbook.Connections("Query - fixtures").OLEDBConnection.Refresh
' Sub ListObjs_RefreshFirstMethod4() -> ActiveWorkbook.Connections("Query - fixtures").Refresh
' Sub ListObjs_RefreshFirstMethod5() -> ActiveWorkbook.QUERIES.Item("fixtures").Refresh *** WINDOWS ONLY ***

Sub ListObjs_RefreshFirstMethod1()
    
    ' Works on Mac
    ' Works in Windows
    
    Debug.Print "--- ListObjs_RefreshFirstMethod1 STARTS ---"
    Debug.Print "OS: " & Chr(34) & Application.OperatingSystem & Chr(34) & "."
    
    Dim msg As String
    
    If ActiveSheet.ListObjects.Count = 0 Then
    
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "ActiveSheet.ListObjects.Count = 0 :0("
        
    Else
    
        StartTime = Now
        
        ActiveSheet.ListObjects(1).Refresh
        
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "The following command was executed:" & vbCr & "ActiveSheet.ListObjects(1).Refresh" & vbCr & vbCr
        msg = msg & "Time taken:" & vbCr & Format(Now - StartTime, "hh:mm:ss")
    
    End If

    
    
    Debug.Print msg
    
    Debug.Print "--- ListObjs_RefreshFirstMethod1 ENDS ---" & vbCr
    
    MsgBox msg

End Sub 'ListObjs_RefreshFirstMethod1

Sub ListObjs_RefreshFirstMethod2()

    ' Works on Mac
    ' Works in Windows
    
    Debug.Print "--- ListObjs_RefreshFirstMethod2 STARTS ---"
    Debug.Print "OS: " & Chr(34) & Application.OperatingSystem & Chr(34) & "."
    
    Dim msg As String
    
    If ActiveSheet.ListObjects.Count = 0 Then
    
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "ActiveSheet.ListObjects.Count = 0 :0("
        
    Else
    
        StartTime = Now
        
        ActiveSheet.ListObjects(1).QueryTable.Refresh
        
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "The following command was executed:" & vbCr & "ActiveSheet.ListObjects(1).QueryTable.Refresh" & vbCr & vbCr
        msg = msg & "Time taken:" & vbCr & Format(Now - StartTime, "hh:mm:ss")
    
    End If

    Debug.Print msg
    
    Debug.Print "--- ListObjs_RefreshFirstMethod2 ENDS ---" & vbCr
    
    MsgBox msg

End Sub ' ListObjs_RefreshFirstMethod2

Sub ListObjs_RefreshFirstMethod3()

    ' Works on Mac
    ' Works in Windows
    
    Debug.Print "--- ListObjs_RefreshFirstMethod3 STARTS ---    "
    Debug.Print "OS: " & Chr(34) & Application.OperatingSystem & Chr(34) & "."
    
    Dim msg As String
    Dim QryName As String
    
    If ActiveSheet.ListObjects.Count = 0 Then
    
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "ActiveSheet.ListObjects.Count = 0 :0("
        
    Else
    
        StartTime = Now
        
        QryName = "Query - " & ActiveSheet.ListObjects(1).Name
        
        ActiveWorkbook.Connections(QryName).OLEDBConnection.Refresh
        
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "The following command was executed:" & vbCr & "ActiveWorkbook.Connections(" & Chr(34) & QryName & Chr(34) & ").OLEDBConnection.Refresh" & vbCr & vbCr
        msg = msg & "Time taken:" & vbCr & Format(Now - StartTime, "hh:mm:ss")
    
    End If

    Debug.Print msg
    
    Debug.Print "--- ListObjs_RefreshFirstMethod3 ENDS ---" & vbCr
    
    MsgBox msg

End Sub ' ListObjs_RefreshFirstMethod3

Sub ListObjs_RefreshFirstMethod4()

    ' Works on Mac
    ' Works in Windows

    Debug.Print "--- ListObjs_RefreshFirstMethod4 STARTS ---"
    Debug.Print "OS: " & Chr(34) & Application.OperatingSystem & Chr(34) & "."
    
    Dim msg As String
    Dim QryName As String
    
    If ActiveSheet.ListObjects.Count = 0 Then
    
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "ActiveSheet.ListObjects.Count = 0 :0("
        
    Else
    
        StartTime = Now
        
        QryName = "Query - " & ActiveSheet.ListObjects(1).Name
        
        ActiveWorkbook.Connections(QryName).Refresh
        
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "The following command was executed:" & vbCr & "ActiveWorkbook.Connections(" & Chr(34) & QryName & Chr(34) & ").Refresh" & vbCr & vbCr
        msg = msg & "Time taken:" & vbCr & Format(Now - StartTime, "hh:mm:ss")
    
    End If

    Debug.Print msg
    
    Debug.Print "--- ListObjs_RefreshFirstMethod4 ENDS ---" & vbCr
    
    MsgBox msg

End Sub ' ListObjs_RefreshFirstMethod4

Sub ListObjs_RefreshFirstMethod5()

    ' Doesn't work on Mac :0(
    ' Works in Windows

    On Error GoTo TheEnd
    
    Debug.Print "--- ListObjs_RefreshFirstMethod5 STARTS ---"
    Debug.Print "OS: " & Chr(34) & Application.OperatingSystem & Chr(34) & "."
    
    Dim msg As String
    Dim NameQry As String
    Dim NameLO As String
    Dim NameUsed As String
    
    If ActiveSheet.ListObjects.Count = 0 Then
    
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "ActiveSheet.ListObjects.Count = 0 :0("
        
    Else
    
        StartTime = Now
        
        NameLO = ActiveSheet.ListObjects(1).Name
        NameQry = "Query - " & NameLO
        NameUsed = NameLO
        ' NameUsed = NameQry -> Not this one
        Debug.Print "NameUsed: " & Chr(34) & NameUsed & Chr(34) & "."
        
        ActiveWorkbook.QUERIES.Item(NameUsed).Refresh
        
        msg = msg & "ActiveSheet.Name:" & vbCr & ActiveSheet.Name & vbCr & vbCr
        msg = msg & "The following command was executed:" & vbCr & "ActiveWorkbook.QUERIES.Item(" & Chr(34) & NameUsed & Chr(34) & ").Refresh" & vbCr & vbCr
        msg = msg & "Time taken:" & vbCr & Format(Now - StartTime, "hh:mm:ss")
    
    End If
    
TheEnd:

    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number:" & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    End If

    Debug.Print msg
    
    Debug.Print "--- ListObjs_RefreshFirstMethod5 ENDS ---" & vbCr
    
    MsgBox msg

End Sub ' ListObjs_RefreshFirstMethod5
