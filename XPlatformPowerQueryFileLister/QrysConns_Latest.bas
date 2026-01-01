Attribute VB_Name = "QrysConns_Latest"
' Check out the following at some point:
' https://community.fabric.microsoft.com/t5/Power-Query/Querytable-refresh-vs-ListObject-Refresh/td-p/2890348
' Sub QrysConns_ListQrysBasic()
' Sub QrysConns_ListConnsDetailed()
' Sub QrysConns_BackgroundRefreshFalseAll()
' Private Sub QrysConns_ExportQrysWin() > Assigned to a button so set to Private
' Private Sub QrysConns_ExportQrysMac() > Assigned to a button so set to Private

Sub QrysConns_ListQrysBasic()

    ' Lists the names of all QUERIES in the ActiveWorkbook & assigns each an index number
    
    Dim msg, Qrys As String
    Dim Qry As WorkbookQuery
    Dim q, pad As Long
    
    pad = Len(CStr(ActiveWorkbook.QUERIES.Count))
    
    If ActiveWorkbook.QUERIES.Count = 0 Then
    
        msg = "ActiveWorkbook.QUERIES.Count = 0"
    
    Else
    
        For Each Qry In ActiveWorkbook.QUERIES
        
            q = q + 1
            
            If Qrys = "" Then
            
                Qrys = Format(q, WorksheetFunction.Rept("0", pad)) & ": " & Qry.Name
            
            Else
            
                Qrys = Qrys & vbCr & Format(q, WorksheetFunction.Rept("0", pad)) & ": " & Qry.Name
            
            End If
        
        Next Qry
        
        msg = msg & "ActiveWorkbook.Name:" & vbCr & Chr(34) & ActiveWorkbook.Name & Chr(34) & "." & vbCr & vbCr
        msg = msg & "--- " & CStr(ActiveWorkbook.QUERIES.Count) & " QUERIES ---" & vbCr & Qrys
    
    End If
    
    Debug.Print msg
    
    ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
    MsgBox msg, vbOKOnly, "Name of each QUERY in the ActiveWorkbook."

End Sub ' QrysConns_ListQrysBasic

Sub QrysConns_ListConnsDetailed()

    Dim msg, ConnType, bgref As String
    Dim Conn As WorkbookConnection
    Dim pad, c, BGRefreshTrue, BGRefreshFalse, BGRefreshNA As Integer
    BGRefreshTrue = 0: BGRefreshFalse = 0: BGRefreshNA = 0
    pad = Len(CStr(ActiveWorkbook.Connections.Count))
    
    msg = "ActiveWorkbook.Name:" & vbCr & ActiveWorkbook.Name
    
    If ActiveWorkbook.Connections.Count = 0 Then
    
        msg = msg & vbCr & vbCr & "ActiveWorkbook.Connections.Count = 0"
    
    Else ' ActiveWorkbook.Connections.Count > 0
    
        msg = msg & vbCr & vbCr & "ActiveWorkbook.Connections.Count = " & CStr(ActiveWorkbook.Connections.Count)
        msg = msg & vbCr & vbCr & "INDEX | TYPE | BACKGROUND REFRESH | NAME"
        
        For Each Conn In ActiveWorkbook.Connections
        
            c = c + 1
            
            msg = msg & vbCr & Format(c, WorksheetFunction.Rept("0", pad))
            
            Select Case Conn.Type
                
                Case 1: ConnType = "OLEDB" ' xlConnectionTypeOLEDB | The initials OLEDB stand for Object Linking and Embedding, Database
                Case 2: ConnType = "ODBC" ' xlConnectionTypeODBC | The initials ODBC stand for Open Database Connectivity
                Case 3: ConnType = "XML MAP" ' xlConnectionTypeXMLMAP
                Case 4: ConnType = "Text" ' xlConnectionTypeTEXT
                Case 5: ConnType = "Web" ' xlConnectionTypeWEB
                Case 6: ConnType = "Data Feed" ' xlConnectionTypeDATAFEED
                Case 7: ConnType = "PowerPivot Model" ' xlConnectionTypeMODEL
                Case 8: ConnType = "Worksheet" ' xlConnectionTypeWORKSHEET
                Case 9: ConnType = "No source" ' xlConnectionTypeNOSOURCE
                Case Else: ConnType = "Else"
            
            End Select
            
            msg = msg & " | " & ConnType
            
            ' The BackgroundQuery property is applicable to certain types of CONNECTIONS, not QUERIES
            
            Select Case ConnType
            
                Case "OLEDB"
                
                    If Conn.OLEDBConnection.BackgroundQuery Then
                    
                        BGRefreshTrue = BGRefreshTrue + 1
                        
                    Else
                    
                        BGRefreshFalse = BGRefreshFalse + 1
                        
                    End If
                    
                    bgref = CStr(Conn.OLEDBConnection.BackgroundQuery)
                
                Case "ODBC"
                
                    If Conn.ODBCConnection.BackgroundQuery Then
                    
                        BGRefreshTrue = BGRefreshTrue + 1
                        
                    Else
                    
                        BGRefreshFalse = BGRefreshFalse + 1
                        
                    End If
                
                    bgref = CStr(Conn.ODBCConnection.BackgroundQuery)
                
                Case Else
            
                    BGRefreshNA = BGRefreshNA + 1
                    
                    bgref = "-----"
            
            End Select
            
            msg = msg & " | " & bgref
            
            msg = msg & " | " & Conn.Name
        
        Next Conn
        
        msg = msg & vbCr & vbCr & "--- Background refreshes ---"
        msg = msg & vbCr & "True :0( " & CStr(BGRefreshTrue)
        msg = msg & vbCr & "False :0) " & CStr(BGRefreshFalse)
        msg = msg & vbCr & "N/A: " & CStr(BGRefreshNA)
    
    End If ' ActiveWorkbook.Connections.Count > 0
    
    Debug.Print "--- len(msg): " & CStr(Len(msg)) & " ---"
    Debug.Print msg
    
    MsgBox msg, vbOKOnly, "Details of each CONNECTION in the ActiveWorkbook."

End Sub ' QrysConns_ListConnsDetailed

Sub QrysConns_BackgroundRefreshFalseAll()

    ' Coverted to ActiveWorkbook
    
    Dim Conn As WorkbookConnection
    Dim msg As String
    Dim ConnCount, WereTrue As Long
    
    For Each Conn In ActiveWorkbook.Connections
    
        Select Case Conn.Type
        
            Case 1 ' OLEDB
            
                ConnCount = ConnCount + 1
                
                If Conn.OLEDBConnection.BackgroundQuery = True Then
                
                    WereTrue = WereTrue + 1
                    
                    Conn.OLEDBConnection.BackgroundQuery = False
                
                End If
            
            Case 2 ' ODBC
            
                ConnCount = ConnCount + 1
                
                If Conn.ODBCConnection.BackgroundQuery = True Then
                
                    WereTrue = WereTrue + 1
                    
                    Conn.ODBCConnection.BackgroundQuery = False
                
                End If
            
            Case Else
            
                ' Do nothing
        
        End Select
    
    Next Conn

    msg = msg & "BackgroundQuery set to False for "
    msg = msg & CStr(WereTrue) & " out of " & CStr(ConnCount) & " connections "
    msg = msg & "in ActiveWorkbook: " & ActiveWorkbook.Name & "."
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' QrysConns_BackgroundRefreshFalseAll

Private Sub QrysConns_ExportQrysWin()

    ' Converted to ActiveWorkbook
    
    Dim objFso As Object
    Dim PathToDesktop As String
    Dim q As Integer
    
    If Left(ActiveWorkbook.Path, 5) = "https" Then
    
        MsgBox "Download " & Chr(34) & ActiveWorkbook.Name & Chr(34) & " to your local machine & try again."
        
    Else
    
        If Right(Environ("USERPROFILE"), 1) = Application.PathSeparator Then
        
            PathToDesktop = Environ("USERPROFILE") & "Desktop" & Application.PathSeparator
        
        Else
        
            PathToDesktop = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator
        
        End If
        
        Set objFso = CreateObject("Scripting.FileSystemObject")
        
        Set myfile = objFso.CreateTextFile( _
            Filename:=PathToDesktop & "Qrys.txt", _
            overwrite:=True, _
            Unicode:=True _
        )
        myfile.writeline WorksheetFunction.Rept("-", 100)
        myfile.writeline "ActiveWorkbook.Path: " & Chr(34) & ActiveWorkbook.Path & Chr(34) & "."
        myfile.writeline "ActiveWorkbook.Name: " & Chr(34) & ActiveWorkbook.Name & Chr(34) & "."
        myfile.writeline "ActiveWorkbook.Queries.Count: " & CStr(ActiveWorkbook.QUERIES.Count) & "."
        myfile.writeline WorksheetFunction.Rept("-", 100)
        myfile.writeline ""
        
        If ActiveWorkbook.QUERIES.Count = 0 Then
        
            myfile.writeline "No queries in this workbook."
        
        Else ' ActiveWorkbook.QUERIES.Count > 0
            
            For q = 1 To ActiveWorkbook.QUERIES.Count
            
                myfile.writeline "********** " & ActiveWorkbook.QUERIES(q).Name & " **********"
                myfile.writeline ActiveWorkbook.QUERIES.Item(q).Formula
                myfile.writeline ""
            
            Next q
        
        End If ' ActiveWorkbook.QUERIES.Count > 0
        
        myfile.Close
        
        Set objFso = Nothing
        
        ActiveWorkbook.FollowHyperlink PathToDesktop & "Qrys.txt"
    
    End If

End Sub ' QrysConns_ExportQrysWin

Private Sub QrysConns_ExportQrysMac() ' Windows OK too so long as you comment out: AppleScriptTask

    ' Coverted to ActiveWorkbook
    
    On Error GoTo TheEnd
    
    ' I have a vague memory of Macs not liking variables declared all on one line like this:
    ' Dim ErrMsg, PathToDesktop As String
    ' ie Declare each variable separately (can't hurt)
    Dim ErrMsg As String
    Dim PathToDesktop As String
    Dim q As Integer
    
    If Left(ActiveWorkbook.Path, 5) = "https" Then
    
        ErrMsg = "Download " & Chr(34) & ActiveWorkbook.Name & Chr(34) & " to your local machine & try again."
        
    Else ' Left(ActiveWorkbook.Path, 5) is not "https"
        
        If Left(Application.OperatingSystem, 3) = "Win" Then
        
            If Right(Environ("USERPROFILE"), 1) = Application.PathSeparator Then
            
                PathToDesktop = Environ("USERPROFILE") & "Desktop" & Application.PathSeparator
            
            Else
            
                PathToDesktop = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator
            
            End If
        
        Else
            
            PathToDesktop = "/Users/" & Environ("USER") & "/Desktop/"
        
        End If
        
        myfile = PathToDesktop & "Qrys.txt"
        
        fileNum = FreeFile(): Open myfile For Output As #fileNum
        
        Print #fileNum, WorksheetFunction.Rept("-", 100)
        Print #fileNum, "ActiveWorkbook.Path: " & Chr(34) & ActiveWorkbook.Path & Chr(34) & "."
        Print #fileNum, "ActiveWorkbook.Name: " & Chr(34) & ActiveWorkbook.Name & Chr(34) & "."
        Print #fileNum, "ActiveWorkbook.Queries.Count: " & CStr(ActiveWorkbook.QUERIES.Count) & "."
        Print #fileNum, WorksheetFunction.Rept("-", 100)
        Print #fileNum, ""
        
        If ActiveWorkbook.QUERIES.Count = 0 Then
        
            Print #fileNum, "No queries in this workbook."
        
        Else ' ActiveWorkbook.QUERIES.Count > 0
            
            For q = 1 To ActiveWorkbook.QUERIES.Count
            
                Print #fileNum, "********** " & ActiveWorkbook.QUERIES(q).Name & " **********"
                Print #fileNum, ActiveWorkbook.QUERIES.Item(q).Formula
                Print #fileNum, ""
            
            Next q
        
        End If ' ActiveWorkbook.QUERIES.Count > 0
        
        Close #fileNum
        
        If Left(Application.OperatingSystem, 3) = "Win" Then
        
            ActiveWorkbook.FollowHyperlink PathToDesktop & "Qrys.txt"
            
        Else
        
            ' --- START ---
            
            PathToHandler = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/"
            Handler = "ShellExScript.scpt"
            Dim HandlerFound As Byte
            
            strFile = Dir(PathToHandler)
            
            Do While Len(strFile) > 0
                
                If strFile = Handler Then
                
                    HandlerFound = 1: Exit Do
                
                End If
                
                strFile = Dir
            
            Loop
            
            If HandlerFound = 0 Then
            
                ErrMsg = "The file: " & Chr(34) & Handler & Chr(34) & " was not found."
            
            Else ' HandlerFound = 1
            
                ' Open containing folder in finder
                
                CommandToRun = "open " & myfile
                
                Debug.Print "CommandToRun:" & vbCr & CommandToRun
                
                CommandResult = AppleScriptTask("ShellExScript.scpt", "ShellEx", CommandToRun)
            
            End If ' HandlerFound = 1
            
            ' --- END ---
        
        End If
    
    End If ' Left(ActiveWorkbook.Path, 5) is not "https"

TheEnd:

    If Err.Number > 0 Then
    
        msg = "--- Error ---" & vbCr & vbCr & "Number:" & CStr(Err.Number) & vbCr & "Desc: " & Err.Description
        
        Err.Clear
    
    Else
    
        If ErrMsg <> "" Then
        
            msg = ErrMsg
        
        Else
        
            msg = ""
        
        End If
    
    End If

    If msg <> "" Then
    
        Debug.Print msg
        
        MsgBox msg
        
    End If

End Sub ' QrysConns_ExportQrysMac


