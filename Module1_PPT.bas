Attribute VB_Name = "Module1"
Option Explicit

Sub ImportABunch()
    'Import one final chart image per slide.
    
    Dim strTemp As String
    Dim strPath As String
    Dim strFileSpec As String
    Dim oSld As Slide
    Dim oPic As Shape, i As Integer
    Dim myWatermark As PowerPoint.Shape
    
    'Path to final chart images
    strPath = "C:\Users\anush\OneDrive\Documents\Sem3\Explorative Information Visualization\Project\ResultCharts\"
    strFileSpec = "img*.svg"
    
    strTemp = Dir(strPath & strFileSpec)
    
    i = 3
    
    Do While strTemp <> ""
        Set oSld = ActivePresentation.Slides.Add(Index:=i, Layout:=ppLayoutBlank)
        'Set oSld = ActivePresentation.Slides(i)
        Set oPic = oSld.Shapes.AddPicture(FileName:=strPath & strTemp, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, _
        Top:=0, _
        Width:=960, _
        Height:=550)
        
        Set myWatermark = oSld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                Left:=865, Top:=525, Width:=200, Height:=50)

        With myWatermark.TextFrame.TextRange
            .Text = "Anusha Porwal"
            With .Font
                .Size = 10
                .Name = "Arial"
                .Color = RGB(0, 0, 0)
            End With
        End With
        
        myWatermark.ZOrder msoBringToFront
    
        i = i + 1
    
    
        With oPic
            .LockAspectRatio = msoFalse
            .ZOrder msoSendToBack
        End With
    
        strTemp = Dir
    Loop

End Sub


