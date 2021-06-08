Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.AutoFilter

Namespace FilterCellsByCellColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_3.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create an auto filter in the sheet and specify the range to be filterd
			sheet.AutoFilters.Range = sheet.Range("G1:G19")

			'Get the column to be filterd
			Dim filtercolumn As FilterColumn = CType(sheet.AutoFilters(0), FilterColumn)

			'Add a color filter to filter the column based on cell color
			sheet.AutoFilters.AddFillColorFilter(filtercolumn, Color.Red)

			'Filter the data.
			sheet.AutoFilters.Filter()

			Dim result As String = "Result-FilterCellsByCellColor.xlsx"

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
