Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ExplodedDoughnut
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbbok
			Dim workbook As New Workbook()

			'Get the first sheet and set its name
			Dim sheet As Worksheet = workbook.Worksheets(0)
			sheet.Name = "ExplodedDoughnut"

			'Set chart data
			CreateChartData(sheet)

			'Add a chart
			Dim chart As Chart = sheet.Charts.Add()
			chart.ChartType = ExcelChartType.DoughnutExploded

			'Set position of chart
			chart.LeftColumn = 1
			chart.TopRow = 6
			chart.RightColumn = 11
			chart.BottomRow = 29

			'Set region of chart data
			chart.DataRange = sheet.Range("A1:B5")
			chart.SeriesDataFromRange = False

			'Chart title
			chart.ChartTitle = "Sales market by country"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12

			For Each cs As ChartSerie In chart.Series
				cs.Format.Options.IsVaryColor = True
				cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
			Next cs

			chart.PlotArea.Fill.Visible = False
			chart.Legend.Position = LegendPositionType.Top

			'Save and Launch
			workbook.SaveToFile("Sample.xlsx",ExcelVersion.Version2010)
			ExcelDocViewer("Sample.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
			'Set value of specified cell
			sheet.Range("A1").Value = "Country"
			sheet.Range("A2").Value = "Cuba"
			sheet.Range("A3").Value = "Mexico"
			sheet.Range("A4").Value = "France"
			sheet.Range("A5").Value = "German"


			sheet.Range("B1").Value = "Sales"
			sheet.Range("B2").NumberValue = 6000
			sheet.Range("B3").NumberValue = 8000
			sheet.Range("B4").NumberValue = 9000
			sheet.Range("B5").NumberValue = 8500

			'Style
			sheet.Range("A1:B1").RowHeight = 15
			sheet.Range("A1:B1").Style.Color = Color.DarkGray
			sheet.Range("A1:B1").Style.Font.Color = Color.White
			sheet.Range("A1:B1").Style.VerticalAlignment = VerticalAlignType.Center
			sheet.Range("A1:B1").Style.HorizontalAlignment = HorizontalAlignType.Center

			sheet.Range("B2:B5").Style.NumberFormat = """$""#,##0"
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
