Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace AutoFitBasedOnCellValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

			'Get first worksheet of the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Set value for B8
			Dim cell As CellRange = worksheet.Range("B8")
			cell.Text = "Welcome to Spire.XLS!"

			'Set the cell style
			Dim style As CellStyle = cell.Style
			style.Font.Size = 16
			style.Font.IsBold = True

			'Autofit column width and row height based on cell value
			cell.AutoFitColumns()
			cell.AutoFitRows()

			'String for output file 
			Dim outputFile As String = "Output.xlsx"

			'Save the file
			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
