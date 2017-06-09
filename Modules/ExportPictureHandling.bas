Attribute VB_Name = "ExportPictureHandling"
'Source Adapted from: http://www.experts-exchange.com/Software/Office_Productivity/Office_Suites/MS_Office/Excel/Q_27403520.html
'May have originated from other cites:
'   http://www.jpsoftwaretech.com/blog/2009/04/export-excel-range-to-a-picture-file-redux/, or
'   http://peltiertech.com/WordPress/export-chart-as-image-file/

Public Function export(shp As Object) As String
Dim pic_object As Shape
Dim pic_height As Double
Dim pic_with As Double
Dim fName As String
Dim tmp_object As Chart

'Exports Shape/Image/Chart to a JPG file by first pasting it into a temporary chart object, then exporting that object as a JPG file

    Set pic_object = Workbooks(shp.Parent.Parent.Name).Sheets(shp.Parent.Name).Shapes(shp.Name)
    
    fName = ActiveWorkbook.path & "\" & pic_object.Name & ".png"
    
    pic_height = pic_object.Height
    pic_width = pic_object.Width
    
    pic_object.CopyPicture appearance:=xlScreen, Format:=xlPicture
    
    Set tmp_object = ActiveSheet.ChartObjects.Add(1, 1, pic_object.Width, pic_object.Height).Chart
    With tmp_object
        .ChartArea.Border.LineStyle = 0
        .Paste
        .export Filename:=fName
        .Parent.Delete
    End With

    export = fName
End Function
