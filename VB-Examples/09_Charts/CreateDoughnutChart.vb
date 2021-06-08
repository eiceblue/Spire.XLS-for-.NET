Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CreateDoughnutChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Insert data
			sheet.Range("A1").Value = "Country"
			sheet.Range("A1").Style.Font.IsBold = True
			sheet.Range("A2").Value = "Cuba"
			sheet.Range("A3").Value = "Mexico"
			sheet.Range("A4").Value = "France"
			sheet.Range("A5").Value = "German"
			sheet.Range("B1").Value = "Sales"
			sheet.Range("B1").Style.Font.IsBold = True
			sheet.Range("B2").NumberValue = 6000
			sheet.Range("B3").NumberValue = 8000
			sheet.Range("B4").NumberValue = 9000
			sheet.Range("B5").NumberValue = 8500

			'Add a new chart, set chart type as doughnut
			Dim chart As Chart = sheet.Charts.Add()
			chart.ChartType = ExcelChartType.Doughnut
			chart.DataRange = sheet.Range("A1:B5")
			chart.SeriesDataFromRange = False

			'Set position of chart
			chart.LeftColumn = 4
			chart.TopRow = 2
			chart.RightColumn = 12
			chart.BottomRow = 22

			'Chart title
			chart.ChartTitle = "Market share by country"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12

			For Each cs As ChartSerie In chart.Series
				cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = True
			Next cs

			chart.Legend.Position = LegendPositionType.Top

			'Save the document
			Dim output As String = "CreateDoughnutChart.xlsx"
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
