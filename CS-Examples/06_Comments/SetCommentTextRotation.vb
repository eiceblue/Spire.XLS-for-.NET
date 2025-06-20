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
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CellValues.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new font object for formatting
            Dim font As ExcelFont = workbook.CreateFont()
            font.FontName = "Arial"
            font.Size = 11
            font.KnownColor = ExcelColors.Orange

            ' Get a specific cell range (E1 in this case)
            Dim range As CellRange = sheet.Range("E1")

            ' Add a comment to the cell
            range.Comment.Text = "This is a comment"

            ' Apply the font formatting to the comment text
            range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font)

            ' Set the vertical alignment of the comment
            range.Comment.VAlignment = CommentVAlignType.Center

            ' Set the horizontal alignment of the comment
            range.Comment.HAlignment = CommentHAlignType.Right

            ' Set the text rotation of the comment
            range.Comment.TextRotation = TextRotationType.LeftToRight

            ' Specify the output file name
            Dim outputFile As String = "Output.xlsx"

            ' Save the workbook to the specified output file
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
