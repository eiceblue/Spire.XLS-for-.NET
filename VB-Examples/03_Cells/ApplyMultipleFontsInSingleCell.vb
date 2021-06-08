Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ApplyMultipleFontsInSingleCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create a font object in workbook, setting the font color, size and type.
			Dim font1 As ExcelFont = workbook.CreateFont()
			font1.KnownColor = ExcelColors.LightBlue
			font1.IsBold = True
			font1.Size = 10

			'Create another font object specifying its properties.
			Dim font2 As ExcelFont = workbook.CreateFont()
			font2.KnownColor = ExcelColors.Red
			font2.IsBold = True
			font2.IsItalic = True
			font2.FontName = "Times New Roman"
			font2.Size = 11

			'Write a RichText string to the cell 'A1', and set the font for it.
			Dim richText As RichText = sheet.Range("H5").RichText
			richText.Text = "This document was created with Spire.XLS for .NET."
			richText.SetFont(0, 29, font1)
			richText.SetFont(31, 48, font2)

			Dim result As String = "Result-ApplyMultipleFontsInSingleCell.xlsx"

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
