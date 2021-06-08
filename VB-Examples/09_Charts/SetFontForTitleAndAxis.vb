Imports Spire.Xls

Namespace SetFontForTitleAndAxis

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\ChartSample1.xlsx")

			'Set font for chart title and chart axis
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			Dim chart As Chart = worksheet.Charts(0)

			'Format the font for the chart title
			chart.ChartTitleArea.Color = Color.Blue
			chart.ChartTitleArea.Size = 20.0
			chart.ChartTitleArea.FontName = "Arial"

			'Format the font for the chart Axis
			chart.PrimaryValueAxis.Font.Color = Color.Gold
			chart.PrimaryValueAxis.Font.Size = 10.0

			chart.PrimaryCategoryAxis.Font.FontName = "Arial"
			chart.PrimaryCategoryAxis.Font.Color = Color.Red
			chart.PrimaryCategoryAxis.Font.Size = 20.0


			'Save the document
			Dim output As String = "SetFontForTitleAndAxis.xlsx"
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
