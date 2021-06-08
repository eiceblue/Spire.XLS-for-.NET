Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyWithOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample.xlsx")

			'Get the first worksheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			'Add a new worksheet as destination sheet
			Dim destinationSheet As Worksheet = workbook.Worksheets.Add("DestSheet")

			'Specify a copy range of original sheet
			Dim cellRange As CellRange = sheet1.Range("B2:D4")

			'Copy the specified range to added worksheet and keep original styles and update reference
			workbook.Worksheets(0).Copy(cellRange, workbook.Worksheets(1), 2, 1, True, True)

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
