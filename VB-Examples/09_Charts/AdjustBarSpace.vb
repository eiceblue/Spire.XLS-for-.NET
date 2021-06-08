Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace AdjustBarSpace

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

			'Get the first worksheet from workbook and then get the first chart from the worksheet
			Dim ws As Worksheet = workbook.Worksheets(0)
			Dim chart As Chart = ws.Charts(0)

			'Adjust the space between bars
			For Each cs As ChartSerie In chart.Series
				cs.Format.Options.GapWidth = 200
				cs.Format.Options.Overlap = 0
			Next cs

			'Save the document
			Dim output As String = "AjustBarSpace.xlsx"
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
