Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ApplySoftEdgesEffect

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample3.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the chart
			Dim chart As Chart = sheet.Charts(0)

			'Specify the size of the soft edge. Value can be set from 0 to 100
			chart.ChartArea.Shadow.SoftEdge = 25

			'Save the document
			Dim output As String = "ApplySoftEdgesEffect.xlsx"
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
