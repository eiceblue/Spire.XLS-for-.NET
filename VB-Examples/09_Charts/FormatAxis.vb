Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace FormatAxis

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
			sheet.Name = "FormatAxis"

			'Set chart data
			CreateChartData(sheet)

			'Add a chart
			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)
			chart.DataRange = sheet.Range("B1:B9")
			chart.SeriesDataFromRange = False
			chart.PlotArea.Visible = False
			chart.TopRow = 10
			chart.BottomRow = 28
			chart.LeftColumn = 2
			chart.RightColumn = 10
			chart.ChartTitle = "Chart with Customized Axis"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12
			Dim cs1 As Spire.Xls.Charts.ChartSerie = chart.Series(0)
			cs1.CategoryLabels = sheet.Range("A2:A9")

			'Format axis
			chart.PrimaryValueAxis.MajorUnit = 8
			chart.PrimaryValueAxis.MinorUnit = 2
			chart.PrimaryValueAxis.MaxValue = 50
			chart.PrimaryValueAxis.MinValue = 0
			chart.PrimaryValueAxis.IsReverseOrder = False
			chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside
			chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside
			chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis
			chart.PrimaryValueAxis.CrossesAt = 0

			'Set NumberFormat
			chart.PrimaryValueAxis.NumberFormat = "$#,##0"
			chart.PrimaryValueAxis.IsSourceLinked = False

			Dim serie As ChartSerie = chart.Series(0)

			For Each dataPoint As ChartDataPoint In serie.DataPoints
				'Format Series
				dataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor
				dataPoint.DataFormat.Fill.ForeColor = Color.LightGreen

				'Set transparency
				dataPoint.DataFormat.Fill.Transparency =0.3
			Next dataPoint

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("Output.xlsx")
		End Sub
		Private Sub CreateChartData(ByVal sheet As Worksheet)
			'Set value of specified cell
			sheet.Range("A1").Value = "Month"
			sheet.Range("A2").Value = "Jan"
			sheet.Range("A3").Value = "Feb"
			sheet.Range("A4").Value = "Mar"
			sheet.Range("A5").Value = "Apr"
			sheet.Range("A6").Value = "May"
			sheet.Range("A7").Value = "Jun"
			sheet.Range("A8").Value = "Jul"
			sheet.Range("A9").Value = "Aug"

			sheet.Range("B1").Value = "Planned"
			sheet.Range("B2").NumberValue = 38
			sheet.Range("B3").NumberValue = 47
			sheet.Range("B4").NumberValue = 39
			sheet.Range("B5").NumberValue = 36
			sheet.Range("B6").NumberValue = 27
			sheet.Range("B7").NumberValue = 25
			sheet.Range("B8").NumberValue = 36
			sheet.Range("B9").NumberValue = 48
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
