Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace HighlightDuplicateUniqueValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
			Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs.AddRange(sheet.Range("C2:C10"))
			Dim format1 As IConditionalFormat = xcfs.AddCondition()
			format1.FormatType = ConditionalFormatType.DuplicateValues
			format1.BackColor = Color.IndianRed

			'Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
			Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs1.AddRange(sheet.Range("C2:C10"))
			Dim format2 As IConditionalFormat = xcfs.AddCondition()
			format2.FormatType = ConditionalFormatType.UniqueValues
			format2.BackColor = Color.Yellow

			Dim result As String = "Result-HighlightDuplicateAndUniqueValues.xlsx"

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
