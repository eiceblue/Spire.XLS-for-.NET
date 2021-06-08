Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SetCommentFillColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()

			'Get the default first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create Excel font
			Dim font As ExcelFont = workbook.CreateFont()
			font.FontName = "Arial"
			font.Size = 11
			font.KnownColor = ExcelColors.Orange

			'Add the comment
			Dim range As CellRange = sheet.Range("A1")
			range.Comment.Text = "This is a comment"
			range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font)

			'Set comment Color
			range.Comment.Fill.FillType = ShapeFillType.SolidColor
			range.Comment.Fill.ForeColor = Color.SkyBlue

			range.Comment.Visible = True

			'String for output file 
			Dim result As String = "SetCommentFillColor_out.xlsx"

			'Save the file
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
