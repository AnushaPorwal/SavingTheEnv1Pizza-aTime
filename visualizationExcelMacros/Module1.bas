Attribute VB_Name = "Module1"
Option Explicit

'This file contains the code to generate the pie charts (for each year and each country) based on the data in the pivot tables.
'The size of the piecharts depends on the totalFoodProduced_t column.


Sub subSavePizzas()
    'loops from year 1979 to modGlobal.Run_to ; For each year:
    '   calls subSetSlicer by passing the year (to change the year in the slicer on the 'work' sheet
    '   calls subSaveImages to save the 3 pie charts generated for the given year
    
    Dim i As Integer, s
    For i = 1979 To Run_to
        Call subSetSlicer(i)
        
        s = Timer + 3
        Do While Timer < s
            DoEvents
        Loop
        
        Call subSaveImages

    Next i
        
End Sub

Sub subSaveImages()
    Dim ws As Worksheet, strFolder As String, strFile As String, strYear As String
    Dim myFileName As String, dSize As Double, dProduce As Double
    
    Set ws = ThisWorkbook.Worksheets("Work")
    strYear = ws.Cells(1, "F")
    strFolder = "C:\Users\anush\OneDrive\Documents\Sem3\Explorative Information Visualization\Project\TempChartFolder\"
    
    Dim objChrt As ChartObject
    Dim myChart As Chart

    'Egypt
    'creating piechart
    Set objChrt = ws.ChartObjects("chtEgypt")
    Set myChart = objChrt.Chart
    
    'Gets the size of the pie chart based on the produce value. Gets size from the function fnGetImgSize.
    dProduce = ws.Cells(22, "F")
    dSize = fnGetImgSize(dProduce)
    
    'sets the size of the piechart
    ws.Shapes("ChtEgypt").Height = dSize
    ws.Shapes("ChtEgypt").Width = dSize
    
    'create the filename
    myFileName = strFolder & "imgE" & CStr(strYear) & ".svg"
    
    'If filename already exits, it overwrites it, else creates a new file
    On Error Resume Next
    Kill myFileName
    On Error GoTo 0
    
    'saves file
    myChart.Export Filename:=myFileName, FilterName:="SVG"
    
    
    'India
    Set objChrt = ws.ChartObjects("chtIndia")
    Set myChart = objChrt.Chart

    dProduce = ws.Cells(23, "F")
    dSize = fnGetImgSize(dProduce)
    
    ws.Shapes("chtIndia").Height = dSize
    ws.Shapes("chtIndia").Width = dSize
    
    myFileName = strFolder & "imgI" & CStr(strYear) & ".svg"

    On Error Resume Next
    Kill myFileName
    On Error GoTo 0

    myChart.Export Filename:=myFileName, FilterName:="SVG"


    'USA
    Set objChrt = ws.ChartObjects("chtUSA")
    Set myChart = objChrt.Chart

    dProduce = ws.Cells(24, "F")
    dSize = fnGetImgSize(dProduce)
    
    ws.Shapes("chtUSA").Height = dSize
    ws.Shapes("chtUSA").Width = dSize

    myFileName = strFolder & "imgU" & CStr(strYear) & ".svg"

    On Error Resume Next
    Kill myFileName
    On Error GoTo 0

    myChart.Export Filename:=myFileName, FilterName:="SVG"



End Sub

Sub subSetSlicer(strYear As Integer)
Attribute subSetSlicer.VB_ProcData.VB_Invoke_Func = " \n14"
    'In the 'work' sheet there is a slicer to set the year.
    'This sub takes strYear as input and changes the slicer value to strYear.
    'This change helps to filter the data in the 3 pivot tables.
    
    Dim i As Integer
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Work")
    
    
    With wb.SlicerCaches("Slicer_year")
        .SlicerItems(CStr(strYear)).Selected = True
        
        For i = 1979 To Run_to
            If i <> strYear Then 'if i not equal to strYear
                .SlicerItems(CStr(i)).Selected = False
            End If
        Next i
        

    End With
End Sub

Function fnGetImgSize(x As Double) As Double
    'This function takes the value of amount of food produced (in tonnes usually), and gets the size that the piechart should be.
    '168 is the maximum amount of food produced (maxValue from column totalFoodProduced_t in 'finalData2use-Copy' sheet)
    '27 is minimum amount of food produced.
    'a and b are the min and max image sizes in cms. It is set in modGlobal as public global variables. You can chage this is you want in modGlobal.
    'multiplying by 28.34646 to convert from cm to points
    
    fnGetImgSize = ((b - a) * x / 168 + (195 * a - 27 * b) / 168) * 28.34646
End Function

