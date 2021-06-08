Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CreateMultiLevelChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Write data to cells
			sheet.Range("A1").Text = "Main Category"
			sheet.Range("A2").Text = "Fruit"
			sheet.Range("A6").Text = "Vegies"
			sheet.Range("B1").Text = "Sub Category"
			sheet.Range("B2").Text = "Bananas"
			sheet.Range("B3").Text = "Oranges"
			sheet.Range("B4").Text = "Pears"
			sheet.Range("B5").Text = "Grapes"
			sheet.Range("B6").Text = "Carrots"
			sheet.Range("B7").Text = "Potatoes"
			sheet.Range("B8").Text = "Celery"
			sheet.Range("B9").Text = "Onions"
			sheet.Range("C1").Text = "Value"
			sheet.Range("C2").Value = "52"
			sheet.Range("C3").Value = "65"
			sheet.Range("C4").Value = "50"
			sheet.Range("C5").Value = "45"
			sheet.Range("C6").Value = "64"
			sheet.Range("C7").Value = "62"
			sheet.Range("C8").Value = "89"
			sheet.Range("C9").Value = "57"

			'//Vertically merge cells from A2 to A5, A6 to A9
			sheet.Range("A2:A5").Merge()
			sheet.Range("A6:A9").Merge()
			sheet.AutoFitColumn(1)
			sheet.AutoFitColumn(2)

			'Add a clustered bar chart to worksheet
			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.BarClustered)
			chart.ChartTitle = "Value"
			chart.PlotArea.Fill.FillType = ShapeFillType.NoFill
			chart.Legend.Delete()
			chart.LeftColumn = 5
			chart.TopRow = 1
			chart.RightColumn = 14

			'Set the data source of series data
			chart.DataRange = sheet.Range("C2:C9")
			chart.SeriesDataFromRange = False
			'Set the data source of category labels
			Dim serie As ChartSerie = chart.Series(0)
			serie.CategoryLabels = sheet.Range("A2:B9")
			'Show multi-level category labels
			chart.PrimaryCategoryAxis.MultiLevelLable = True

			'Save the document
			Dim output As String = "CreateMultiLevelChart.xlsx"
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
