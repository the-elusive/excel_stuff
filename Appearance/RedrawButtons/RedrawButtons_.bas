Attribute VB_Name = "RedrawButtons_"
Private Sub RedrawButtons()

    ' Ordinary rather than ActiveX buttons
    
    ' Edit myArray as required
    Dim myArray, msg, ButtonRng, ButtonTxt, ButtonAct As String
    'eg myArray = myArray & "I16:I17|Open folder|DoOpenContainingFolder"
    myArray = myArray & "G1:I1|Refresh|QrysConns_RefreshFirstListObj" & vbCr
    myArray = myArray & "J1:K1|Calibri 11|FirstListObjCalibriEleven" & vbCr
    myArray = myArray & "L1|Icons|FormIconsColourEach"
    
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



