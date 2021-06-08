Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetDBNumFormatting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			workbook.CreateEmptySheets(1)

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set value for cells
			sheet.Range("A1").Value2 = 123
			sheet.Range("A2").Value2 = 456
			sheet.Range("A3").Value2 = 789

			'Get the cell range
			Dim range As CellRange = sheet.Range("A1:A3")

			'Set the DB num format
			range.NumberFormat = "[DBNum2][$-804]General"

			'Auto fit columns
			range.AutoFitColumns()

			'Save the document
			Dim output As String = "SetDBNumFormatting_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
