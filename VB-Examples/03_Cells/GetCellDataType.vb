Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace GetCellDataType
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_2.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the cell types of the cells in range "C13:F13"
			For Each range As CellRange In sheet.Range("H2:H7")
				Dim cellType As XlsWorksheet.TRangeValueType = sheet.GetCellType(range.Row, range.Column, False)
				sheet(range.Row, range.Column + 1).Text = cellType.ToString()
				sheet(range.Row, range.Column + 1).Style.Font.Color = Color.Red
				sheet(range.Row, range.Column + 1).Style.Font.IsBold = True
			Next range

			Dim result As String = "Result-GetCellDataType.xlsx"

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
