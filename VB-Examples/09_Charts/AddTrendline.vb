Imports Spire.Xls

Namespace AddTrendline

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample2.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'select chart and set logarithmic trendline
			Dim chart As Chart = sheet.Charts(0)
			chart.ChartTitle = "Logarithmic Trendline"
			chart.Series(0).TrendLines.Add(TrendLineType.Logarithmic)
			'select chart and set moving_average trendline
			Dim chart1 As Chart = sheet.Charts(1)
			chart1.ChartTitle = "Moving Average Trendline"
			chart1.Series(0).TrendLines.Add(TrendLineType.Moving_Average)
			'select chart and set linear trendline
			Dim chart2 As Chart = sheet.Charts(2)
			chart2.ChartTitle = "Linear Trendline"
			chart2.Series(0).TrendLines.Add(TrendLineType.Linear)
			'select chart and set exponential trendline
			Dim chart3 As Chart = sheet.Charts(3)
			chart3.ChartTitle = "Exponential Trendline"
			chart3.Series(0).TrendLines.Add(TrendLineType.Exponential)

			'Save the document
			Dim output As String = "AddTrendline.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
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
