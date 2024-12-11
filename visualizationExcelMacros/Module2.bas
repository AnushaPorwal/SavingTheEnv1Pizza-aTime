Attribute VB_Name = "Module2"
Sub subMakeMainChart()
    'Code loops through all years and generates the Main chart with a timeline+slider to see which year the chart is for.
    'Then it saves the chart and timeline+slider together as an image.
    
    Dim i As Integer, s
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Work")
    
    'Aligns the placement of the timeline and slider with the left of the chart.
    ws.Shapes("timeline").Left = ws.ChartObjects("chtMain").Left
    ws.Shapes("timeline").Width = ws.ChartObjects("chtMain").Width
    ws.Shapes("slider").Left = ws.Shapes("timeline").Left
    
    'loops through all years from 1979 to modGlobal.Run_to
    For i = 1979 To Run_to
        'sets the year in the slicer on 'work' sheet to be able to filter out the data in the 3 pivot tables.
        Call subSetSlicer(i)
        
        s = Timer + 2
        Do While Timer < s
            DoEvents
        Loop
        
        'Imports the piecharts into the chart for corresponding country and year.
        Call subImportImg("E", i)
        Call subImportImg("I", i)
        Call subImportImg("U", i)
        
        'sets the position of the slider on the timeline, based on the current year (i)
        Call subSetSlider
        
        s = Timer + 2
        Do While Timer < s
            DoEvents
        Loop
        
        'saves the chart and timeline+slider as a combined image
        Call ExportAsSVG


    Next i
        
End Sub

Sub subImportImg(strInitial As String, iYear As Integer)
Attribute subImportImg.VB_ProcData.VB_Invoke_Func = " \n14"
    'subInitial is the country code (for which the piechart needs to be imported into chart)
    'iYear is the year for which the chart is being created
    
    Dim strFile As String, iPoint As Integer
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Work")
    
    'iPoint is a reference to the scatter plot point in chtMain
    Select Case strInitial
        Case "E":
            iPoint = 1
        Case "I":
            iPoint = 2
        Case "U":
            iPoint = 3
    End Select
    
    'location of where the piecharts are saved - from where they are being imported.
    strFile = "C:\Users\anush\OneDrive\Documents\Sem3\Explorative Information Visualization\Project\TempChartFolder\img" & strInitial & CStr(iYear) & ".svg"
    Dim cht As Chart
    
    ws.Activate
    Set cht = ws.ChartObjects("chtMain").Chart
    
    'imports the correct piechart
    cht.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(iPoint).Select
    Selection.MarkerStyle = -4147
    With Selection.Format.Fill
        .Visible = msoTrue
        .UserPicture strFile
    End With
    
    'Removes the box border around the scatter plot point in the chart.
    cht.FullSeriesCollection(1).Select
    Selection.Format.Line.Visible = msoFalse
    
End Sub

Sub subSetSlider()
    'based on current year (of slicer or in cell F1), this sub sets the position of the slider on the timeline bar (which is placed on top of chtMain).
    
    Dim ws As Worksheet, strFile As String, iYear As Integer
    Dim myFileName As String, dSize As Double, dProduce As Double
    Dim timelineLeft As Double, timelineRight As Double, slider
    
    
    Set ws = ThisWorkbook.Worksheets("Work")
    iYear = ws.Cells(1, "F") 'Get current year
    
    
    Dim sTimeline As Shape, sSlider As Shape
    Set sTimeline = ws.Shapes("timeline")
    Set sSlider = ws.Shapes("slider")
    
    'sets left position of the slider based on current year and range(1979, modGlobal.Run_to)
    sSlider.Left = sTimeline.Left + (iYear - 1979) * (sTimeline.Width - sSlider.Width) / (Run_to - 1979)


End Sub

Sub subSaveMainChart()
    'Is not in use anymore
    'Was previously used to save chtMain into folders.
    'This was before the timeline+slider was added
    
    Dim ws As Worksheet, strFolder As String, strFile As String, strYear As String
    Dim myFileName As String, dSize As Double, dProduce As Double
    
    Set ws = ThisWorkbook.Worksheets("Work")
    strYear = ws.Cells(1, "F")
    strFolder = "C:\Users\anush\OneDrive\Documents\Sem3\Explorative Information Visualization\Project\ResultCharts\"
    
    Dim objChrt As ChartObject
    Dim myChart As Chart
    Dim groupShape As Shape, shpRes As Shape

    'Set objChrt = ws.ChartObjects("chtMain")
    'Set myChart = objChrt.Chart
    
    'ActiveSheet.Shapes.Range(Array("slider", "timeline", "chtMain")).Select
    Set groupShape = ws.Shapes.Range(Array("slider", "timeline"))
    groupShape.Select
    
    'Set groupShape = ws.Shapes.Range("Group 7")
    'groupShape.Select
    
    'Set shpRes = ws.Shapes(ws.Shapes.Count)

    myFileName = strFolder & "img" & CStr(strYear) & ".svg"

    On Error Resume Next
    Kill myFileName
    On Error GoTo 0

    groupShape.Export Filename:=myFileName, FilterName:="SVG"
End Sub

Sub SaveChartAndShapeAsImage()
    'This function is not in use!!
    'Attempts to do the following:
    'Grouping Chart and Shape gave an entity of type Object. Issue was that Object type didn't have export method. So was unable to save it as an image.
    'This method copies the chart and the sape and combines them into a temporary shape in the 'Temp' sheet.
    'After this it does some reformatting and now it can be saved as an image (that has both chart and shape)
    
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim shapeObj As Shape
    Dim combinedRange As Range
    Dim imageFilePath As String
    Dim tempShape As Shape

    'Set your worksheet and objects here
    Set ws = ThisWorkbook.Worksheets("Work")    ' Replace with your sheet name
    Set chartObj = ws.ChartObjects("chtMain")   ' Adjust the index or name of the chart
    Set shapeObj = ws.Shapes("slider")          ' Adjust the index or name of the shape

    'Combine the chart and shape into a temporary shape
    chartObj.Copy
    Set combinedRange = ws.Range("A100")        ' A temporary location to place the combined object
    combinedRange.PasteSpecial
    Set tempShape = ws.Shapes(ws.Shapes.Count)

    'Move the shape and resize the temp shape to match combined content
    With tempShape
        .Top = Application.Min(chartObj.Top, shapeObj.Top)
        .Left = Application.Min(chartObj.Left, shapeObj.Left)
        .Width = Application.Max(chartObj.Left + chartObj.Width, shapeObj.Left + shapeObj.Width) - .Left
        .Height = Application.Max(chartObj.Top + chartObj.Height, shapeObj.Top + shapeObj.Height) - .Top
    End With

    'Save as a picture
    imageFilePath = ThisWorkbook.Path & "\CombinedImage.png" ' Set your desired path and file name
    tempShape.Chart.Export Filename:=imageFilePath

    'Delete the temporary shape
    tempShape.Delete

    ' Notify the user
    MsgBox "Image saved successfully to: " & imageFilePath
End Sub

Sub ExportAsSVG()
    'Grouping Chart and Shape gave an entity of type Object. Issue was that Object type didn't have export method. So was unable to save it as an image.
    'This method copies the chart and the sape and combines them into a temporary shape in the 'Temp' sheet.
    'After this it does some reformatting and now it can be saved as an image (that has both chart and shape)
    
    Dim wsh As Worksheet, wsT As Worksheet
    Dim shp As Shape
    Dim fil As Variant
    Dim cho As ChartObject
    Dim myFileName As String, strFolder As String, strFile As String, iYear As Integer
    Dim ts1 As Shape, ts2 As Shape, ts3 As Shape
    Dim groupShape As Shape
    
    Set wsh = ThisWorkbook.Worksheets("Work")
    Set wsT = ThisWorkbook.Worksheets("Temp")
    
    iYear = wsh.Cells(1, "F")  'current year
    strFolder = "C:\Users\anush\OneDrive\Documents\Sem3\Explorative Information Visualization\Project\ResultCharts\"
    myFileName = strFolder & "img" & CStr(iYear) & ".svg"

    Dim dLeftSlider As Double, dWidthTimeline As Double, dWidthSlider As Double
    
    wsT.Activate
    
    Set shp = wsh.Shapes("timeline")
    
    shp.CopyPicture
    wsT.Range("A4").PasteSpecial
    Set ts1 = wsT.Shapes(wsT.Shapes.Count)
    ts1.Top = ts1.Top - 8
    dWidthTimeline = ts1.Width
    
    
    Set shp = wsh.Shapes("slider")
    shp.CopyPicture
    wsT.Activate
    s = Timer + 1
    Do While Timer < s
        DoEvents
    Loop
    
    wsT.Range("A4").PasteSpecial
    Set ts2 = wsT.Shapes(wsT.Shapes.Count)
    dWidthSlider = ts2.Width
    ts2.Left = (iYear - 1979) * (dWidthTimeline - dWidthSlider) / (Run_to - 1979)
    ts2.Top = ts2.Top - 14
    
    Set cho = wsh.ChartObjects("chtMain")
    cho.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    wsT.Activate
    wsT.Range("A6").PasteSpecial
    Set ts3 = wsT.Shapes(wsT.Shapes.Count)
    
    s = Timer + 1
    Do While Timer < s
        DoEvents
    Loop
    
    Set groupShape = wsT.Shapes.Range(Array(ts1.Name, ts2.Name, ts3.Name)).Group

    'Create a ChartObject to host the grouped shape for exporting
    Dim tempChart As ChartObject
    Set tempChart = wsT.ChartObjects.Add(Left:=groupShape.Left, Top:=groupShape.Top, Width:=groupShape.Width, Height:=groupShape.Height)

    'Copy the grouped shape into the chart
    groupShape.Copy
    
    s = Timer + 1
    Do While Timer < s
        DoEvents
    Loop
    
    tempChart.Chart.Paste
    
    tempChart.Chart.Export Filename:=myFileName, FilterName:="SVG"
    
    s = Timer + 1
    Do While Timer < s
        DoEvents
    Loop
    
    tempChart.Delete
    groupShape.Delete
    
End Sub
