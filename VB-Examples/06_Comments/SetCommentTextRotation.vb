Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Drawing.Imaging
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.Collections

Namespace SetCommentTextRotation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CellValues.xlsx")

			'Get the default first  worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create Excel font
			Dim font As ExcelFont = workbook.CreateFont()
			font.FontName = "Arial"
			font.Size = 11
			font.KnownColor = ExcelColors.Orange

			'Add the comment
			Dim range As CellRange = sheet.Range("E1")
			range.Comment.Text = "This is a comment"
			range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font)

			' Set its vertical and horizontal alignment 
			range.Comment.VAlignment = CommentVAlignType.Center
			range.Comment.HAlignment = CommentHAlignType.Right

			'Set the comment text rotation
			range.Comment.TextRotation = TextRotationType.LeftToRight

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
