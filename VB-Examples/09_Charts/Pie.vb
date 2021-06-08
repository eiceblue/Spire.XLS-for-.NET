Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace Pie
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
			sheet.Name = "Pie Chart"

			'Add a chart
			Dim chart As Chart = Nothing
			If checkBox1.Checked Then
				chart = sheet.Charts.Add(ExcelChartType.Pie3D)
			Else
				chart = sheet.Charts.Add(ExcelChartType.Pie)
			End If

			'Set chart data
			CreateChartData(sheet)

			'Set region of chart data
			chart.DataRange = sheet.Range("B2:B5")
			chart.SeriesDataFromRange = False

			'Set position of chart
			chart.LeftColumn = 1
			chart.TopRow = 6
			chart.RightColumn = 9
			chart.BottomRow = 25

			'Chart title
			chart.ChartTitle = "Sales by year"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12

			Dim cs As ChartSerie = chart.Series(0)
			cs.CategoryLabels = sheet.Range("A2:A5")
			cs.Values = sheet.Range("B2:B5")
			cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True

			chart.PlotArea.Fill.Visible = False

			'Save and Launch
			workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010)
			ExcelDocViewer("Output.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
			'Set value of specified cell
			sheet.Range("A1").Value = "Year"
			sheet.Range("A2").Value = "2002"
			sheet.Range("A3").Value = "2003"
			sheet.Range("A4").Value = "2004"
			sheet.Range("A5").Value = "2005"

			sheet.Range("B1").Value = "Sales"
			sheet.Range("B2").NumberValue = 4000
			sheet.Range("B3").NumberValue = 6000
			sheet.Range("B4").NumberValue = 7000
			sheet.Range("B5").NumberValue = 8500

			'Style
			sheet.Range("A1:B1").RowHeight = 15
			sheet.Range("A1:B1").Style.Color = Color.DarkGray
			sheet.Range("A1:B1").Style.Font.Color = Color.White
			sheet.Range("A1:B1").Style.VerticalAlignment = VerticalAlignType.Center
			sheet.Range("A1:B1").Style.HorizontalAlignment = HorizontalAlignType.Center

			sheet.Range("B2:C5").Style.NumberFormat = """$""#,##0"
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
