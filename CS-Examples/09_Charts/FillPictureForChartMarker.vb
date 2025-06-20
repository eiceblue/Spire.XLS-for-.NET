Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace FillPictureForChartMarker
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Define the input Excel file path
            Dim inputFile As String = "..\..\..\..\..\..\Data\FillChartMarker.xlsx"

            ' Define the image file path for marker fill
            Dim imageFile As String = "..\..\..\..\..\..\Data\E-iceblueLogo.png"

            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified input file path
            workbook.LoadFromFile(inputFile)

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = worksheet.Charts(0)

            ' Set the color of the line properties for the first series in the chart to yellow
            chart.Series(0).Format.LineProperties.Color = Color.Yellow

            ' Set the marker style for the first series in the chart to picture
            chart.Series(0).Format.MarkerStyle = ChartMarkerType.Picture

            ' Get the marker fill of the first series
            Dim markerFill1 As IShapeFill = chart.Series(0).DataFormat.MarkerFill

            ' Set a custom picture as the marker fill for the first series using the specified image file
            markerFill1.CustomPicture(imageFile)

            ' Get the marker fill of the second series
            Dim markerFill2 As IShapeFill = chart.Series(1).DataFormat.MarkerFill

            ' Set the color of the line properties for the second series in the chart to red
            chart.Series(1).Format.LineProperties.Color = Color.Red

            ' Set the texture of the marker fill for the second series to granite
            markerFill2.Texture = GradientTextureType.Granite

            ' Set the color of the line properties for the first series in the chart to blue
            chart.Series(0).Format.LineProperties.Color = Color.Blue

            ' Get the marker fill of the third series
            Dim markerFill3 As IShapeFill = chart.Series(2).DataFormat.MarkerFill

            ' Set the pattern of the marker fill for the third series to 10% gradient
            markerFill3.Pattern = GradientPatternType.Pat10Percent

            ' Set the foreground color of the marker fill for the third series to light gray
            markerFill3.ForeColor = Color.LightGray

            ' Set the background color of the marker fill for the third series to orange
            markerFill3.BackColor = Color.Orange

            ' Save the modified workbook to a new file named "result.xlsx" using Excel 2013 version
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()
            'View the document
            FileViewer("result.xlsx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
