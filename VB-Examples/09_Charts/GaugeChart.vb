Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace GaugeChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
			Dim workbook As New Workbook()

			'Get the first sheet and set its name
			Dim sheet As Worksheet = workbook.Worksheets(0)
			sheet.Name = "Gauge Chart"

			'Set chart data
			CreateChartData(sheet)

			'Add a Doughnut chart
			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.Doughnut)
			chart.DataRange = sheet.Range("A1:A5")
			chart.SeriesDataFromRange = False
			chart.HasLegend = True

			'Set the position of chart
			chart.LeftColumn = 2
			chart.TopRow = 7
			chart.RightColumn = 9
			chart.BottomRow = 25

			'Get the series 1
			Dim cs1 As ChartSerie = CType(chart.Series("Value"), ChartSerie)
			cs1.Format.Options.DoughnutHoleSize = 60
			cs1.DataFormat.Options.FirstSliceAngle = 270

			'Set the fill color
			cs1.DataPoints(0).DataFormat.Fill.ForeColor = Color.Yellow
			cs1.DataPoints(1).DataFormat.Fill.ForeColor = Color.PaleVioletRed
			cs1.DataPoints(2).DataFormat.Fill.ForeColor = Color.DarkViolet
			cs1.DataPoints(3).DataFormat.Fill.Visible = False

			'Add a series with pie chart
			Dim cs2 As ChartSerie = CType(chart.Series.Add("Pointer", ExcelChartType.Pie), ChartSerie)

			'Set the value
			cs2.Values = sheet.Range("D2:D4")
			cs2.UsePrimaryAxis = False
			cs2.DataPoints(0).DataLabels.HasValue = True
			cs2.DataFormat.Options.FirstSliceAngle = 270
			cs2.DataPoints(0).DataFormat.Fill.Visible = False
			cs2.DataPoints(1).DataFormat.Fill.FillType = ShapeFillType.SolidColor
			cs2.DataPoints(1).DataFormat.Fill.ForeColor = Color.Black
			cs2.DataPoints(2).DataFormat.Fill.Visible = False

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", FileFormat.Version2010)
			ExcelDocViewer("Output.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
			'Set value of specified cell
			sheet.Range("A1").Value = "Value"
			sheet.Range("A2").Value = "30"
			sheet.Range("A3").Value = "60"
			sheet.Range("A4").Value = "90"
			sheet.Range("A5").Value = "180"
			sheet.Range("C2").Value = "value"
			sheet.Range("C3").Value = "pointer"
			sheet.Range("C4").Value = "End"
			sheet.Range("D2").Value = "10"
			sheet.Range("D3").Value = "1"
			sheet.Range("D4").Value = "189"
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
