Imports Spire.Xls

Namespace ShowOrHideGridLine

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample2.xlsx")

			'Get the first and second worksheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)
			Dim sheet2 As Worksheet = workbook.Worksheets(1)

			'Hide grid line in the first worksheet
			sheet1.GridLinesVisible = False
			'Show grid line in the first worksheet
			sheet2.GridLinesVisible = True

			'Save the document
			Dim output As String = "ShowOrHideGridLine.xlsx"
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
