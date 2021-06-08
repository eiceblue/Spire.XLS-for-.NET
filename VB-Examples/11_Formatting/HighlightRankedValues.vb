Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace HighlightRankedValues
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

			'Apply conditional formatting to range "D2:D10" to highlight the top 2 values.
			Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs.AddRange(sheet.Range("D2:D10"))
			Dim format1 As IConditionalFormat = xcfs.AddTopBottomCondition(TopBottomType.Top, 2)
			format1.FormatType = ConditionalFormatType.TopBottom
			format1.BackColor = Color.Red

			'Apply conditional formatting to range "E2:E10" to highlight the bottom 2 values.
			Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs1.AddRange(sheet.Range("E2:E10"))
			Dim format2 As IConditionalFormat = xcfs1.AddTopBottomCondition(TopBottomType.Bottom,2)
			format2.FormatType = ConditionalFormatType.TopBottom
			format2.BackColor = Color.ForestGreen

			Dim result As String = "Result-HighlightTopAndBottomRankedValues.xlsx"

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
