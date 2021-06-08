Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace RemoveRowBasedOnKeyword
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorkbookToHTML.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Find the string
			Dim cr As CellRange = sheet.FindString("Address", False, False)

			'Delete the row which includes the string
			sheet.DeleteRow(cr.Row)

			'Save to file
			workbook.SaveToFile("RemoveRowBasedOnKeyword.xlsx", ExcelVersion.Version2010)

			'View the document
			FileViewer("RemoveRowBasedOnKeyword.xlsx")
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
