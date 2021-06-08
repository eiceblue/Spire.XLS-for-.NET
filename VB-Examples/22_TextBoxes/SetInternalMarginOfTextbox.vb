Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace SetInternalMarginOfTextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add a textbox to the sheet and set its position and size.
			Dim textbox As XlsTextBoxShape = TryCast(sheet.TextBoxes.AddTextBox(4, 2, 100, 300), XlsTextBoxShape)

			'Set the text on the textbox.
			textbox.Text = "Insert TextBox in Excel and set the margin for the text"
			textbox.HAlignment = CommentHAlignType.Center
			textbox.VAlignment = CommentVAlignType.Center

			'Set the inner margins of the contents.
			textbox.InnerLeftMargin = 1
			textbox.InnerRightMargin = 3
			textbox.InnerTopMargin = 1
			textbox.InnerBottomMargin = 1

			Dim result As String = "Result-SetInternalMarginOfTextbox.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the MS Excel file.
			ExcelDocViewer(result)
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
