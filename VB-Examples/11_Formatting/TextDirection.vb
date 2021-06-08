Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace TextDirection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Add a new worksheet to the Excel object
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Access the "B5" cell from the worksheet
			Dim cell As CellRange = sheet.Range("B5")

			'Add some value to the "B5" cell
			cell.Text = "Hello Spire!"

			'Set the reading order from right to left of the text in the "B5" cell
			cell.Style.ReadingOrder = ReadingOrderType.RightToLeft

			Dim result As String = "TextDirection_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
		   FileViewer(result)
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
