Attribute VB_Name = "ChromedriverVBA"
' Private Sub RedrawButtons()
' Private Sub QrysConns_ExportQrysWin()
' Private Sub SeleniumTester()
' Private Sub RefreshChromedriverLO()
' Private Sub OpenLink()
' Function GetChromeDriverVersion() As String
' Function GetChromeVersion() As String

Private Sub RedrawButtons()

    ' Ordinary rather than ActiveX buttons
    
    Dim msg As String
    Dim ButtonRng As String
    Dim ButtonTxt As String
    Dim ButtonAct As String
    
    ' Edit myArray as required
    Dim myArray As String
    myArray = myArray & "B2:C3|Test Selenium|SeleniumTester" & vbCr
    myArray = myArray & "B4:C5|Open link|OpenLink" & vbCr
    myArray = myArray & "B6:B7|Qrys|QrysConns_ExportQrysWin" & vbCr
    myArray = myArray & "C6:C7|Refresh|RefreshChromedriverLO"
    
    Dim ButtonCount, answer As Integer
    ButtonCount = UBound(Split(myArray, vbCr)) - LBound(Split(myArray, vbCr)) + 1
    
    msg = msg & "About to add " & CStr(ButtonCount) & " buttons to the tab:" & vbCr & Chr(34) & ActiveSheet.Name & Chr(34) & "."
    msg = msg & vbCr & vbCr & "Did you delete the old ones? This procedure doesn't do that automatically."
    
    answer = MsgBox(msg, vbYesNo + vbDefaultButton2)
    
    Select Case answer
    
        Case 6
        
            msg = ""
    
        Case 7
        
            msg = "You clicked " & Chr(34) & "no" & Chr(34) & "."
    
        Case Else
        
            ' You can't cancel
    
    End Select
    
    If msg <> "" Then
    
        MsgBox msg
    
    Else ' msg = ""
    
        For Each Item In Split(myArray, vbCr)
            
            ButtonRng = Split(Item, "|")(0)
            ButtonTxt = Split(Item, "|")(1)
            ButtonAct = Split(Item, "|")(2)
            
            Set theRange = Range(ButtonRng)
            
            ActiveSheet.Buttons.Add( _
                theRange.Left, _
                theRange.Top, _
                theRange.Width, _
                theRange.Height _
            ).Select
            
            With Selection
            
                .Caption = ButtonTxt
                .OnAction = ButtonAct
                ' .VerticalAlignment = xlTop
                
                ' You can also choose to bold the text on certain buttons
                If ButtonTxt = "SCRAPE" Or ButtonTxt = "Format" Or ButtonTxt = "Refresh" Then
                
                    With .Characters(Start:=1, Length:=Len(ButtonTxt)).Font
                    
                        .Name = "Calibri"
                        .FontStyle = "Bold"
                        ' .Size = 8
                    
                    End With
                
                End If
            
            End With
        
        Next Item
        
        [A1].Select ' Deselects the last button you added
    
    End If ' msg = ""
    
End Sub ' RedrawButtons

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
        myfile.writeline "ActiveWorkbook.Queries.Count: " & CStr(ActiveWorkbook.Queries.Count) & "."
        myfile.writeline WorksheetFunction.Rept("-", 100)
        myfile.writeline ""
        
        If ActiveWorkbook.Queries.Count = 0 Then
        
            myfile.writeline "No queries in this workbook."
        
        Else ' ActiveWorkbook.QUERIES.Count > 0
            
            For q = 1 To ActiveWorkbook.Queries.Count
            
                myfile.writeline "********** " & ActiveWorkbook.Queries(q).Name & " **********"
                myfile.writeline ActiveWorkbook.Queries.Item(q).Formula
                myfile.writeline ""
            
            Next q
        
        End If ' ActiveWorkbook.QUERIES.Count > 0
        
        myfile.Close
        
        Set objFso = Nothing
        
        ActiveWorkbook.FollowHyperlink PathToDesktop & "Qrys.txt"
    
    End If

End Sub ' QrysConns_ExportQrysWin

Private Sub SeleniumTester()
    
    On Error GoTo TheEnd
    
    StartTime = Now
    
    Dim msg As String
    Dim PageTitle As String: PageTitle = "Page title missing :0("
    
    GetUrl = "https://www.bbc.co.uk/sport/football/teams/liverpool"
    
    Dim BrowserInstance As New WebDriver ' *** Nice succinct way of doing this ***
    
    With BrowserInstance
    
        .SetPreference "profile.managed_default_content_settings.images", 2 ' 1 for allow 2 for block
        .Start "chrome"
        .Window.Maximize
        .Get GetUrl
    
    End With
    
    PageTitle = BrowserInstance.Title
    
    ' Application.Wait (Now + TimeValue("0:00:05"))
    
    BrowserInstance.Close
    Set BrowserInstance = Nothing
    
TheEnd:

    If Err.Number > 0 Then
    
        msg = "Unuccessful test :0(" & vbCr & vbCr & _
        "Error as follows : " & vbCr & _
        "Err.Number : " & Err.Number & vbCr & _
        "Err.Description : " & Err.Description
        
        Err.Clear
    
    Else
    
        msg = "Successful test :0)" & vbCr & vbCr & _
        "Time taken: " & Format(Now - StartTime, "hh:mm:ss") & "." & vbCr & vbCr & _
        "PageTitle: " & vbCr & PageTitle
    
    End If
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' SeleniumTester

Private Sub RefreshChromedriverLO()

    StartTime = Now
    Dim msg As String
    Dim i As Integer
    Dim ReqLO As String: ReqLO = "Chromedriver"
    Dim FoundLO As Boolean
    
    If ActiveSheet.ListObjects.Count = 0 Then
    
        msg = "No list objects on tab: " & Chr(34) & ActiveSheet.Name & Chr(34) & "."
    
    Else
    
        For i = 1 To ActiveSheet.ListObjects.Count
        
            If ActiveSheet.ListObjects(i).Name = ReqLO Then ' Case sensitive!
            
                FoundLO = True
                
                Exit For
            
            End If
        
        Next i
        
        If FoundLO Then
        
            ActiveSheet.ListObjects(ReqLO).Refresh
            
            Sheets("Chromedriver").Range("LastRefreshed") = _
            "Last refreshed: " & Format(Now, "ddd dd-mmm-yyyy hh:mm:ss") & "."
            
            msg = "Refreshed list object " & Chr(34) & ReqLO & Chr(34) & " " & _
            "on tab " & Chr(34) & ActiveSheet.Name & Chr(34) & "." & vbCr & vbCr & _
            "Time taken: " & Format(Now - StartTime, "hh:mm:ss") & "."
        
        Else ' FoundLO = False
        
            msg = "Did not find the list object " & Chr(34) & ReqLO & Chr(34) & " " & _
            "on tab " & Chr(34) & ActiveSheet.Name & Chr(34) & " :0("
        
        End If
    
    End If
    
    Debug.Print msg
    
    MsgBox msg

End Sub ' RefreshChromedriverLO

Private Sub OpenLink()

    If Left(Selection.Cells(1).Value, 5) <> "https" Then
    
        MsgBox "The left 5 characters " & _
        "in [" & Replace(Selection.Cells(1).Address, "$", "") & "] " & _
        "on tab " & Chr(34) & ActiveSheet.Name & Chr(34) & " " & _
        "aren't " & Chr(34) & "https" & Chr(34) & " :0("
    
    Else
    
        ThisWorkbook.FollowHyperlink Selection.Cells(1).Value
    
    End If

End Sub ' OpenLink

Function GetChromeDriverVersion() As String
    Dim fs As Object
    Dim charPath As String
    
    charPath = "C:\Program Files\SeleniumBasic\chromedriver.exe"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FileExists(charPath) Then
        GetChromeDriverVersion = fs.GetFileVersion(charPath)
    Else
        GetChromeDriverVersion = "ChromeDriver Not Found"
    End If
End Function ' GetChromeDriverVersion

Function GetChromeVersion() As String
    Dim fs As Object
    Dim charPath As String
    
    ' Default path for 64-bit Chrome
    charPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FileExists(charPath) Then
        GetChromeVersion = fs.GetFileVersion(charPath) & " (64-bit)"
    Else
        ' Fallback for 32-bit path if 64-bit doesn't exist
        charPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        If fs.FileExists(charPath) Then
            GetChromeVersion = fs.GetFileVersion(charPath) & " (32-bit)"
        Else
            GetChromeVersion = "Chrome Not Found"
        End If
    End If
End Function ' GetChromeVersion
