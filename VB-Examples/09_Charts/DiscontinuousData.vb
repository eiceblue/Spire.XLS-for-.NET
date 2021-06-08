Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.ComponentModel
Imports System.Text

Namespace DiscontinuousData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a Workbook from disk
			Dim book As New Workbook()
			book.LoadFromFile("..\..\..\..\..\..\Data\DiscontinuousData.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = book.Worksheets(0)

			'Add a chart
			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)
			chart.SeriesDataFromRange = False

			'Set the position of chart
			chart.LeftColumn = 1
			chart.TopRow = 10
			chart.RightColumn = 10
			chart.BottomRow = 24

			'Add a series
			Dim cs1 As ChartSerie = CType(chart.Series.Add(), ChartSerie)

			'Set the name of the cs1
			cs1.Name = sheet.Range("B1").Value

			'Set discontinuous values for cs1
			cs1.CategoryLabels = sheet.Range("A2:A3").AddCombinedRange(sheet.Range("A5:A6")).AddCombinedRange(sheet.Range("A8:A9"))
			cs1.Values = sheet.Range("B2:B3").AddCombinedRange(sheet.Range("B5:B6")).AddCombinedRange(sheet.Range("B8:B9"))

			'Set the chart type
			cs1.SerieType = ExcelChartType.ColumnClustered

			'Add a series
			Dim cs2 As ChartSerie = CType(chart.Series.Add(), ChartSerie)
			cs2.Name = sheet.Range("C1").Value
			cs2.CategoryLabels = sheet.Range("A2:A3").AddCombinedRange(sheet.Range("A5:A6")).AddCombinedRange(sheet.Range("A8:A9"))
			cs2.Values = sheet.Range("C2:C3").AddCombinedRange(sheet.Range("C5:C6")).AddCombinedRange(sheet.Range("C8:C9"))
			cs2.SerieType = ExcelChartType.ColumnClustered

			chart.ChartTitle = "Chart"
			chart.ChartTitleArea.Font.Size = 20
			chart.ChartTitleArea.Color = Color.Black

			chart.PrimaryValueAxis.HasMajorGridLines = False

			'Save and Launch
			book.SaveToFile("Output.xlsx",ExcelVersion.Version2010)
			ExcelDocViewer("Output.xlsx")
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
