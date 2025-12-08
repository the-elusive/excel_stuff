Attribute VB_Name = "SelectionVarTypes"
Sub GotoFirstEmptyCell()

    Dim FirstEmptyCell As String
    
    For Each MyCell In Selection.Cells
    
        Select Case VarType(MyCell)
        
            Case 0
            
                FirstEmptyCell = MyCell.Address
                
                Exit For
            
            Case Else
            
                ' Do nothing
        
        End Select
        
    Next MyCell
    
    If FirstEmptyCell = "" Then
    
        MsgBox "There are no empty cells."
    
    Else
    
        Range(FirstEmptyCell).Select
    
    End If

End Sub ' GotoFirstEmptyCell

Sub GotoFirstTextCell()

    Dim FirstTextCell As String
    
    For Each MyCell In Selection.Cells
    
        Select Case VarType(MyCell)
        
            Case 8
            
                FirstTextCell = MyCell.Address
                
                Exit For
            
            Case Else
            
                ' Do nothing
        
        End Select
        
    Next MyCell
    
    If FirstTextCell = "" Then
    
        MsgBox "There are no text cells."
    
    Else
    
        Range(FirstTextCell).Select
    
    End If

End Sub ' GotoFirstTextCell

Sub GetVarTypesInSelection()

    answer = MsgBox("Show all variable types?", vbYesNo + vbDefaultButton2)
    
    Select Case answer
    
        Case 6
            
            ' MyMessage = "You clicked " & Chr(34) & "yes" & Chr(34) & "."
            
            ShowAllVarTypes = "Y"
    
        Case 7
        
            ' MyMessage = "You clicked " & Chr(34) & "no" & Chr(34) & "."
            
            ShowAllVarTypes = "N"
    
        Case Else
        
            ' You can't cancel
    
    End Select
    
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
    ' cf : https://support.unicomsi.com/manuals/intelligence/75/index.html#page/Server%20User%20Guides/AdvancedExpressions.33.160.html
    
    ' No datatype given for the following variables given that they can have both text & numbers assigned to them
    
    Dim vbEmptys
    Dim vbNulls
    Dim vbIntegers
    Dim vbLongs
    Dim vbSingles
    Dim vbDoubles
    Dim vbCurrencys
    Dim vbDates
    Dim vbStrings
    Dim vbObjects
    Dim vbErrors
    Dim vbBooleans
    Dim vbVariants
    Dim vbDataObjects
    Dim vbDecimals
    Dim vbBytes
    Dim vbLongLongs
    Dim vbUserDefinedTypes
    Dim vbArrays
    Dim vbOthers

    vbEmptys = 0
    vbNulls = 0
    vbIntegers = 0
    vbLongs = 0
    vbSingles = 0
    vbDoubles = 0
    vbCurrencys = 0
    vbDates = 0
    vbStrings = 0
    vbObjects = 0
    vbErrors = 0
    vbBooleans = 0
    vbVariants = 0
    vbDataObjects = 0
    vbDecimals = 0
    vbBytes = 0
    vbLongLongs = 0
    vbUserDefinedTypes = 0
    vbArrays = 0
    vbOthers = 0

    For Each MyCell In Selection.Cells
    
        Select Case VarType(MyCell)
        
            Case 0
                
                vbEmptys = vbEmptys + 1
            
            Case 1

                vbNulls = vbNulls + 1

            Case 2

                vbIntegers = vbIntegers + 1
            
            Case 3
            
                vbLongs = vbLongs + 1

            Case 4

                vbSingles = vbSingles + 1

            Case 5

                vbDoubles = vbDoubles + 1

            Case 6

                vbCurrencys = vbCurrencys + 1

            Case 7

                vbDates = vbDates + 1

            Case 8

                vbStrings = vbStrings + 1
            
            Case 9
                
                vbObjects = vbObjects + 1
            
            Case 10
            
                vbErrors = vbErrors + 1

            Case 11
                
                vbBooleans = vbBooleans + 1
            
            Case 12

                vbVariants = vbVariants + 1
            
            Case 13

                vbDataObjects = vbDataObjects + 1

            Case 14
                
                vbDecimals = vbDecimals + 1
            
            Case 17
            
                vbBytes = vbBytes + 1

            Case 20
            
                vbLongLongs = vbLongLongs + 1
            
            Case 36
                
                vbUserDefinedTypes = vbUserDefinedTypes + 1
            
            Case 8192
                
                vbArrays = vbArrays + 1
            
            Case Else
            
                vbOthers = vbOthers + 1
        
        End Select
    
    Next MyCell
    
    ' Replace nils with "---" so the numbers stand out more STARTS
    
    If vbEmptys = 0 Then
    
        vbEmptys = "---"
        
    End If
    
    If vbNulls = 0 Then
    
        vbNulls = "---"
        
    End If
    
    If vbIntegers = 0 Then
    
        vbIntegers = "---"
        
    End If
    
    If vbLongs = 0 Then
    
        vbLongs = "---"
        
    End If
    
    If vbSingles = 0 Then
    
        vbSingles = "---"
        
    End If
    
    If vbDoubles = 0 Then
    
        vbDoubles = "---"
        
    End If
    
    If vbCurrencys = 0 Then
    
        vbCurrencys = "---"
        
    End If
    
    If vbDates = 0 Then
    
        vbDates = "---"
        
    End If
    
    If vbStrings = 0 Then
    
        vbStrings = "---"
        
    End If
    
    If vbObjects = 0 Then
    
        vbObjects = "---"
        
    End If
    
    If vbErrors = 0 Then
    
        vbErrors = "---"
        
    End If
    
    If vbBooleans = 0 Then
    
        vbBooleans = "---"
    
    End If
    
    If vbVariants = 0 Then
    
        vbVariants = "---"
    
    End If
    
    If vbDataObjects = 0 Then
    
        vbDataObjects = "---"
    
    End If
    
    If vbDecimals = 0 Then
    
        vbDecimals = "---"
    
    End If
    
    If vbBytes = 0 Then
    
        vbBytes = "---"
    
    End If
    
    If vbLongLongs = 0 Then
        
        vbLongLongs = "---"
    
    End If
    
    If vbUserDefinedTypes = 0 Then
    
        vbUserDefinedTypes = "---"
    
    End If
    
    If vbArrays = 0 Then
    
        vbArrays = "---"
    
    End If
    
    If vbOthers = 0 Then
    
        vbOthers = "---"
        
    End If
    
    ' Replace nils with "---" so the numbers stand out more ENDS
    
    ' --- Build concluding message bit by bit STARTS ---
    
    MyMessage = "VarTypes in your selection as follows : " & vbCr & vbCr
    
    ' vbEmptys
    
    If vbEmptys = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbEmptys = " & Format(vbEmptys, "#,##0") & vbCr
        
        End If
        
    Else ' vbEmptys <> "---"
    
        MyMessage = MyMessage & "vbEmptys = " & Format(vbEmptys, "#,##0") & vbCr
    
    End If
    
    ' vbNulls
    
    If vbNulls = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbNulls = " & Format(vbNulls, "#,##0") & vbCr
        
        End If
        
    Else ' vbNulls <> "---"
    
        MyMessage = MyMessage & "vbNulls = " & Format(vbNulls, "#,##0") & vbCr
    
    End If
    
    ' vbIntegers
    
    If vbIntegers = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbIntegers = " & Format(vbIntegers, "#,##0") & vbCr
        
        End If
        
    Else ' vbIntegers <> "---"
    
        MyMessage = MyMessage & "vbIntegers = " & Format(vbIntegers, "#,##0") & vbCr
    
    End If
    
    ' vbLongs
    
    If vbLongs = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbLongs = " & Format(vbLongs, "#,##0") & vbCr
        
        End If
        
    Else ' vbLongs <> "---"
    
        MyMessage = MyMessage & "vbLongs = " & Format(vbLongs, "#,##0") & vbCr
    
    End If
    
    ' vbSingles
    
    If vbSingles = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbSingles = " & Format(vbSingles, "#,##0") & vbCr
        
        End If
        
    Else ' vbSingles <> "---"
    
        MyMessage = MyMessage & "vbSingles = " & Format(vbSingles, "#,##0") & vbCr
    
    End If
    
    ' vbDoubles
    
    If vbDoubles = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbDoubles = " & Format(vbDoubles, "#,##0") & vbCr
        
        End If
        
    Else ' vbDoubles <> "---"
    
        MyMessage = MyMessage & "vbDoubles = " & Format(vbDoubles, "#,##0") & vbCr
    
    End If
    
    ' vbCurrencys
    
    If vbCurrencys = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbCurrencys = " & Format(vbCurrencys, "#,##0") & vbCr
        
        End If
        
    Else ' vbCurrencys <> "---"
    
        MyMessage = MyMessage & "vbCurrencys = " & Format(vbCurrencys, "#,##0") & vbCr
    
    End If
    
    ' vbDates
    
    If vbDates = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbDates = " & Format(vbDates, "#,##0") & vbCr
        
        End If
        
    Else ' vbDates <> "---"
    
        MyMessage = MyMessage & "vbDates = " & Format(vbDates, "#,##0") & vbCr
    
    End If
    
    ' vbStrings
    
    If vbStrings = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbStrings = " & Format(vbStrings, "#,##0") & vbCr
        
        End If
        
    Else ' vbStrings <> "---"
    
        MyMessage = MyMessage & "vbStrings = " & Format(vbStrings, "#,##0") & vbCr
    
    End If
    
    ' vbObjects
    
    If vbObjects = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbObjects = " & Format(vbObjects, "#,##0") & vbCr
        
        End If
        
    Else ' vbObjects <> "---"
    
        MyMessage = MyMessage & "vbObjects = " & Format(vbObjects, "#,##0") & vbCr
    
    End If
    
    ' vbErrors
    
    If vbErrors = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbErrors = " & Format(vbErrors, "#,##0") & vbCr
        
        End If
        
    Else ' vbErrors <> "---"
    
        MyMessage = MyMessage & "vbErrors = " & Format(vbErrors, "#,##0") & vbCr
    
    End If
    
    ' vbBooleans
    
    If vbBooleans = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbBooleans = " & Format(vbBooleans, "#,##0") & vbCr
        
        End If
        
    Else ' vbBooleans <> "---"
    
        MyMessage = MyMessage & "vbBooleans = " & Format(vbBooleans, "#,##0") & vbCr
    
    End If
    
    ' vbVariants
    
    If vbVariants = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbVariants = " & Format(vbVariants, "#,##0") & vbCr
        
        End If
        
    Else ' vbVariants <> "---"
    
        MyMessage = MyMessage & "vbVariants = " & Format(vbVariants, "#,##0") & vbCr
    
    End If
    
    ' vbDataObjects
    
    If vbDataObjects = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbDataObjects = " & Format(vbDataObjects, "#,##0") & vbCr
        
        End If
        
    Else ' vbDataObjects <> "---"
    
        MyMessage = MyMessage & "vbDataObjects = " & Format(vbDataObjects, "#,##0") & vbCr
    
    End If
    
    ' vbDecimals
    
    If vbDecimals = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbDecimals = " & Format(vbDecimals, "#,##0") & vbCr
        
        End If
        
    Else ' vbDecimals <> "---"
    
        MyMessage = MyMessage & "vbDecimals = " & Format(vbDecimals, "#,##0") & vbCr
    
    End If
    
    ' vbBytes
    
    If vbBytes = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbBytes = " & Format(vbBytes, "#,##0") & vbCr
        
        End If
        
    Else ' vbBytes <> "---"
    
        MyMessage = MyMessage & "vbBytes = " & Format(vbBytes, "#,##0") & vbCr
    
    End If
    
    ' vbLongLongs
    
    If vbLongLongs = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbLongLongs = " & Format(vbLongLongs, "#,##0") & vbCr
        
        End If
        
    Else ' vbLongLongs <> "---"
    
        MyMessage = MyMessage & "vbLongLongs = " & Format(vbLongLongs, "#,##0") & vbCr
    
    End If
    
    ' vbUserDefinedTypes
    
    If vbUserDefinedTypes = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbUserDefinedTypes = " & Format(vbUserDefinedTypes, "#,##0") & vbCr
        
        End If
        
    Else ' vbUserDefinedTypes <> "---"
    
        MyMessage = MyMessage & "vbUserDefinedTypes = " & Format(vbUserDefinedTypes, "#,##0") & vbCr
    
    End If
    
    ' vbArrays
    
    If vbArrays = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbArrays = " & Format(vbArrays, "#,##0") & vbCr
        
        End If
        
    Else ' vbArrays <> "---"
    
        MyMessage = MyMessage & "vbArrays = " & Format(vbArrays, "#,##0") & vbCr
    
    End If
    
    ' vbOthers
    
    If vbOthers = "---" Then
    
        If ShowAllVarTypes = "Y" Then
            
            ' Format("---", "#,##0") = "---"
            
            MyMessage = MyMessage & "vbOthers = " & Format(vbOthers, "#,##0") & vbCr
        
        End If
        
    Else ' vbOthers <> "---"
    
        MyMessage = MyMessage & "vbOthers = " & Format(vbOthers, "#,##0") & vbCr
    
    End If
    
    ' --- Build concluding message bit by bit ENDS ---
    
    MsgBox MyMessage
    
    ' MsgBox "VarTypes in your selection as follows : " & vbCr & vbCr & _
    "vbEmptys = " & Format(vbEmptys, "#,##0") & vbCr & _
    "vbNulls = " & Format(vbNulls, "#,##0") & vbCr & _
    "vbIntegers = " & Format(vbIntegers, "#,##0") & vbCr & _
    "vbLongs = " & Format(vbLongs, "#,##0") & vbCr & _
    "vbSingles = " & Format(vbSingles, "#,##0") & vbCr & _
    "vbDoubles = " & Format(vbDoubles, "#,##0") & vbCr & _
    "vbCurrencys = " & Format(vbCurrencys, "#,##0") & vbCr & _
    "vbDates = " & Format(vbDates, "#,##0") & vbCr & _
    "vbStrings = " & Format(vbStrings, "#,##0") & vbCr & _
    "vbObjects = " & Format(vbObjects, "#,##0") & vbCr & _
    "vbErrors = " & Format(vbErrors, "#,##0") & vbCr & _
    "vbBooleans = " & Format(vbBooleans, "#,##0") & vbCr & _
    "vbVariants = " & Format(vbVariants, "#,##0") & vbCr & _
    "vbDataObjects = " & Format(vbDataObjects, "#,##0") & vbCr & _
    "vbDecimals = " & Format(vbDecimals, "#,##0") & vbCr & _
    "vbBytes = " & Format(vbBytes, "#,##0") & vbCr & _
    "vbLongLongs = " & Format(vbLongLongs, "#,##0") & vbCr & _
    "vbUserDefinedTypes = " & Format(vbUserDefinedTypes, "#,##0") & vbCr & _
    "vbArrays = " & Format(vbArrays, "#,##0") & vbCr & _
    "vbOthers = " & Format(vbOthers, "#,##0")

End Sub ' GetVarTypesInSelection



