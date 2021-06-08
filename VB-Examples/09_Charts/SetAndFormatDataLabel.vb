Imports Spire.Xls

Namespace SetAndFormatDataLabel

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			workbook.CreateEmptySheets(1)
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Name = "Demo"
			sheet.Range("A1").Value = "Month"
			sheet.Range("A2").Value = "Jan"
			sheet.Range("A3").Value = "Feb"
			sheet.Range("A4").Value = "Mar"
			sheet.Range("A5").Value = "Apr"
			sheet.Range("A6").Value = "May"
			sheet.Range("A7").Value = "Jun"
			sheet.Range("B1").Value = "Peter"
			sheet.Range("B2").NumberValue = 25
			sheet.Range("B3").NumberValue = 18
			sheet.Range("B4").NumberValue = 8
			sheet.Range("B5").NumberValue = 13
			sheet.Range("B6").NumberValue = 22
			sheet.Range("B7").NumberValue = 28

			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.LineMarkers)
			chart.DataRange = sheet.Range("B1:B7")
			chart.PlotArea.Visible = False
			chart.SeriesDataFromRange = False
			chart.TopRow = 5
			chart.BottomRow = 26
			chart.LeftColumn = 2
			chart.RightColumn = 11
			chart.ChartTitle = "Data Labels Demo"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12
			Dim cs1 As Spire.Xls.Charts.ChartSerie = chart.Series(0)
			cs1.CategoryLabels = sheet.Range("A2:A7")

			cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
			cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = False
			cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = False
			cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
			cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
			cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". "

			cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9
			cs1.DataPoints.DefaultDataPoint.DataLabels.Color = Color.Red
			cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri"
			cs1.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center

			'Save the document
			Dim output As String = "SetAndFormatDataLabel.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the file
			ExcelDocViewer(output)
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
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
