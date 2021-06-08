Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace Line
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
			sheet.Name = "Line Chart"

			'Set chart data
			CreateChartData(sheet)

			'Add a chart
			Dim chart As Chart = sheet.Charts.Add()
			If checkBox1.Checked Then
				chart.ChartType = ExcelChartType.Line3D
			Else
				chart.ChartType = ExcelChartType.Line
			End If

			'Set region of chart data
			chart.DataRange = sheet.Range("A1:E5")

			'Set position of chart
			chart.LeftColumn = 1
			chart.TopRow = 6
			chart.RightColumn = 11
			chart.BottomRow = 29

			'Set chart title
			chart.ChartTitle = "Sales market by country"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12

			chart.PrimaryCategoryAxis.Title = "Month"
			chart.PrimaryCategoryAxis.Font.IsBold = True
			chart.PrimaryCategoryAxis.TitleArea.IsBold = True

			chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
			chart.PrimaryValueAxis.HasMajorGridLines = False
			chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
			chart.PrimaryValueAxis.MinValue = 1000
			chart.PrimaryValueAxis.TitleArea.IsBold = True

			For Each cs As ChartSerie In chart.Series
				cs.Format.Options.IsVaryColor = True
				cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True

				If Not checkBox1.Checked Then
					cs.DataFormat.MarkerStyle = ChartMarkerType.Circle
				End If
			Next cs

			chart.PlotArea.Fill.Visible = False

			chart.Legend.Position = LegendPositionType.Top
			workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010)
			ExcelDocViewer("Output.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
			'Set value of specified cell
			sheet.Range("A1").Value = "Country"
			sheet.Range("A2").Value = "Cuba"
			sheet.Range("A3").Value = "Mexico"
			sheet.Range("A4").Value = "France"
			sheet.Range("A5").Value = "German"


			sheet.Range("B1").Value = "Jun"
			sheet.Range("B2").NumberValue = 3300
			sheet.Range("B3").NumberValue = 2300
			sheet.Range("B4").NumberValue = 4500
			sheet.Range("B5").NumberValue = 6700


			sheet.Range("C1").Value = "Jul"
			sheet.Range("C2").NumberValue = 7500
			sheet.Range("C3").NumberValue = 2900
			sheet.Range("C4").NumberValue = 2300
			sheet.Range("C5").NumberValue = 4200


			sheet.Range("D1").Value = "Aug"
			sheet.Range("D2").NumberValue = 7400
			sheet.Range("D3").NumberValue = 6900
			sheet.Range("D4").NumberValue = 7800
			sheet.Range("D5").NumberValue = 4200


			sheet.Range("E1").Value = "Sep"
			sheet.Range("E2").NumberValue = 8000
			sheet.Range("E3").NumberValue = 7200
			sheet.Range("E4").NumberValue = 8500
			sheet.Range("E5").NumberValue = 5600

			'Style
			sheet.Range("A1:E1").RowHeight = 15
			sheet.Range("A1:E1").Style.Color = Color.DarkGray
			sheet.Range("A1:E1").Style.Font.Color = Color.White
			sheet.Range("A1:E1").Style.VerticalAlignment = VerticalAlignType.Center
			sheet.Range("A1:E1").Style.HorizontalAlignment = HorizontalAlignType.Center

			sheet.Range("B2:D5").Style.NumberFormat = """$""#,##0"
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
