Imports Spire.Xls

Namespace SetGlobalCustomFonts
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Set custom font directory
			Dim fontPath() As String = { "..\..\..\..\..\..\Data\fonts" }

			' Create a new workbook object
			Dim workbook As New Workbook()

			Workbook.SetGlobalCustomFontsFolders(fontPath)

			' Load an existing Excel file from the specified path
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SpecialFont.xlsx")

			' Save the workbook to PDF 
			Dim result As String = "output.pdf"
			workbook.SaveToFile(result, FileFormat.PDF)

			' Dispose of the workbook object
			workbook.Dispose()

			' View the document using a file viewer
			FileViewer(result)

			Me.Close()
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
