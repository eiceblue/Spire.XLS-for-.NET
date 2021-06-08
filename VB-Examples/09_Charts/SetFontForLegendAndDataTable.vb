Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace SetFontForLegendAndDataTable

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
			Dim chart As Chart = ws.Charts(0)

			'Create a font with specified size and color
			Dim font As ExcelFont = workbook.CreateFont()
			font.Size = 14.0
			font.Color = Color.Red

			'Apply the font to chart Legend
			chart.Legend.TextArea.SetFont(font)

			'Apply the font to chart DataLabel
			For Each cs As ChartSerie In chart.Series
				cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
			Next cs

			'Save the document
			Dim output As String = "SetFontForLegendAndDataTable.xlsx"
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
