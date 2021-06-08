Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace CloneExcelFontStyle
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

			'Add the text to the Excel sheet cell range A1.
			sheet.Range("A1").Text = "Text1"

			'Set A1 cell range's CellStyle.
			Dim style As CellStyle = workbook.Styles.Add("style")
			style.Font.FontName = "Calibri"
			style.Font.Color = Color.Red
			style.Font.Size = 12
			style.Font.IsBold = True
			style.Font.IsItalic = True
			sheet.Range("A1").CellStyleName = style.Name

			'Clone the same style for B2 cell range.
			Dim csOrieign As CellStyle = style.clone()
			sheet.Range("B2").Text = "Text2"
			sheet.Range("B2").CellStyleName = csOrieign.Name

			'Clone the same style for C3 cell range and then reset the font color for the text.
			Dim csGreen As CellStyle = style.clone()
			csGreen.Font.Color = Color.Green
			sheet.Range("C3").Text = "Text3"
			sheet.Range("C3").CellStyleName = csGreen.Name

			Dim result As String = "Result-CloneExcelFontStyle.xlsx"

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
