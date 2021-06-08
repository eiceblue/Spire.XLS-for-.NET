Imports Spire.Xls

Namespace LoadAndSaveFileWithMacro

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\MacroSample.xls")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("A5").Text = "This is a simple test!"

			'Save the document
			Dim output As String = "LoadAndSaveFileWithMacro.xls"
			workbook.SaveToFile(output, ExcelVersion.Version97to2003)

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
