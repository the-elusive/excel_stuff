Attribute VB_Name = "LaunchEditors"
' Attribute VB_Name = "LaunchEditors"

Sub AwesomeLink()

    ThisWorkbook.FollowHyperlink "https://www.microsoft.com/en-us/download/details.aspx?id=50745"

End Sub

Sub LaunchVisualBasicEditor()

    Application.CommandBars.ExecuteMso "VisualBasic"

End Sub ' LaunchVisualBasicEditor

Sub LaunchPowerQueryEditor()

    If Left(Application.OperatingSystem, 3) = "Mac" Then

        MsgBox "Procedure " & Chr(34) & "LaunchPowerQueryEditor" & Chr(34) & " doesn't work on a Mac I'm afraid :0("

    Else

        Application.CommandBars.ExecuteMso "PowerQueryLaunchQueryEditor"

    End If

End Sub ' LaunchPowerQueryEditor


