Attribute VB_Name = "ExportSelectionAsImage_"
Sub ExportSelectionAsImage()

    Dim rng As Range
    Dim chtObj As ChartObject
    Dim imgPath As String

    Set rng = Selection

    imgPath = ThisWorkbook.Path & "\SelectionImage.png"

    ' Copy range as picture
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' Create temporary chart
    Set chtObj = ActiveSheet.ChartObjects.Add( _
        Left:=rng.Left, _
        Top:=rng.Top, _
        Width:=rng.Width, _
        Height:=rng.Height)

    With chtObj
        .Activate

        ' Paste picture
        .Chart.Paste

        ' Let Excel finish rendering
        DoEvents

        ' Export image
        .Chart.Export Filename:=imgPath, FilterName:="PNG"

        ' Delete temp chart
        .Delete
    End With

    MsgBox "Exported to:" & vbCrLf & imgPath

End Sub

