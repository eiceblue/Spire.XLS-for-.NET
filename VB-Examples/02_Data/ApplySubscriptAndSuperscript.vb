Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ApplySubscriptAndSuperscript
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("B2").Text = "This is an example of Subscript:"
			sheet.Range("D2").Text = "This is an example of Superscript:"

			'Set the rtf value of "B3" to "R100-0.06".
			Dim range As CellRange = sheet.Range("B3")
			range.RichText.Text = "R100-0.06"

			'Create a font. Set the IsSubscript property of the font to "true".
			Dim font As ExcelFont = workbook.CreateFont()
			font.IsSubscript = True
			font.Color = Color.Green

			'Set font for specified range of the text in "B3".
			range.RichText.SetFont(4, 8, font)

			'Set the rtf value of "D3" to "a2 + b2 = c2".
			range = sheet.Range("D3")
			range.RichText.Text = "a2 + b2 = c2"

			'Create a font. Set the IsSuperscript property of the font to "true".
			font = workbook.CreateFont()
			font.IsSuperscript = True

			'Set font for specified range of the text in "D3".
			range.RichText.SetFont(1, 1, font)
			range.RichText.SetFont(6, 6, font)
			range.RichText.SetFont(11, 11, font)

			sheet.AllocatedRange.AutoFitColumns()

			Dim result As String = "Result-ApplySubscriptAndSuperscript.xlsx"

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
