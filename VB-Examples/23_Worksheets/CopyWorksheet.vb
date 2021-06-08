Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace CopyWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim sourceWorkbook As New Workbook()

			'Load the source Excel document from disk
			sourceWorkbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

			'Get the first worksheet
			Dim srcWorksheet As Worksheet = sourceWorkbook.Worksheets(0)

			'Create a workbook
			Dim targetWorkbook As New Workbook()

			'Load the target Excel document from disk
			targetWorkbook.LoadFromFile("..\..\..\..\..\..\Data\sample.xlsx")

			'Add a new worksheet
			Dim targetWorksheet As Worksheet = targetWorkbook.Worksheets.Add("added")

			'Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
			targetWorksheet.CopyFrom(srcWorksheet)

			'String for output file 
			Dim outputFile As String = "Output.xlsx"

			'Save the file
			targetWorkbook.SaveToFile(outputFile, ExcelVersion.Version2013)

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
