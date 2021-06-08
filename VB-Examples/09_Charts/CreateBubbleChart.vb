Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CreateBubbleChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateBubbleChart.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim chart As Chart = sheet.Charts.Add(ExcelChartType.Bubble)

			'Chart title
			chart.ChartTitle = "Bubble"
			chart.ChartTitleArea.IsBold = True
			chart.ChartTitleArea.Size = 12

			chart.DataRange = sheet.Range("A1:C5")
			chart.SeriesDataFromRange = False

			chart.Series(0).Bubbles = sheet.Range("C2:C5")

			'Set position of chart
			chart.LeftColumn = 7
			chart.TopRow = 6
			chart.RightColumn = 16
			chart.BottomRow = 29

			workbook.SaveToFile("CreateBubbleChart.xlsx", ExcelVersion.Version2010)

			'View the document
			FileViewer("CreateBubbleChart.xlsx")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
