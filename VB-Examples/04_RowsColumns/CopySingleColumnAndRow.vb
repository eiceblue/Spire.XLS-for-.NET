Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopySingleColumnAndRow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

			'Get the first worksheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			'Specify a destination range to copy one column 
			Dim columnCells As CellRange = sheet1.Range("G1:G19")

			'Copy the second column to destination range 
			sheet1.Columns(1).Copy(columnCells)

			'Specify a destination range to copy one row 
			Dim rowCells As CellRange = sheet1.Range("A21:E21")

			'Copy the first row to destination range 
			sheet1.Rows(0).Copy(rowCells)

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
