Imports Spire.Xls

Namespace SetCellFillPattern

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CommonTemplate.xlsx")

			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Set cell color
			worksheet.Range("B7:F7").Style.Color = Color.Yellow
			'Set cell fill pattern
			worksheet.Range("B8:F8").Style.FillPattern = ExcelPatternType.Percent125Gray

			'Save the document
			Dim output As String = "SetCellFillPattern.xlsx"
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
