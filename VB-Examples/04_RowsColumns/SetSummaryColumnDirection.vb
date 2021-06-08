Imports Spire.Xls

Namespace SetSummaryColumnDirection

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

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Group Columns
			sheet.GroupByColumns(1, 4, True)

			'Set summary columns to right of details
			sheet.PageSetup.IsSummaryColumnRight = True

			'Save the document
			Dim output As String = "SetSummaryColumnDirection.xlsx"
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
