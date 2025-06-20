Imports Spire.Xls

Namespace AddSignatureLine
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook instance
			Dim workbook As New Workbook()

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add a signature line 
			sheet.Range("A1").AddSignatureLine("Rose","manager", "manager@test.com", "a short text",False,True)

			'Save the file
			Dim file As String = "AddSignatureLine.xlsx"
			workbook.SaveToFile(file,ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			OutputViewer(file)

		End Sub
		Private Sub OutputViewer(ByVal fileName As String)
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
