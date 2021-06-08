Imports Spire.Xls

Namespace FillChartElementWithPicture

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

			'Get the first worksheet from workbook
			Dim ws As Worksheet = workbook.Worksheets(0)
			'Get the first chart
			Dim chart As Chart = ws.Charts(0)

			' A. Fill chart area with image
			chart.ChartArea.Fill.CustomPicture(Image.FromFile("..\..\..\..\..\..\Data\background.png"), "None")
			chart.PlotArea.Fill.Transparency = 0.9

			'// B.Fill plot area with image
			'chart.PlotArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");

			'Save the document
			Dim output As String = "FillChartElementWithPicture.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2010)

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
