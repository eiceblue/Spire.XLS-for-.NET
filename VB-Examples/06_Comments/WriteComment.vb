Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace WriteComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteComment.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Creates font
			Dim font As ExcelFont=workbook.CreateFont()
			font.FontName="Arial"
			font.Size=11
			font.KnownColor = ExcelColors.Orange
			Dim fontBlue As ExcelFont = workbook.CreateFont()
			fontBlue.KnownColor = ExcelColors.LightBlue
			Dim fontGreen As ExcelFont = workbook.CreateFont()
			fontGreen.KnownColor = ExcelColors.LightGreen

			Dim range As CellRange = sheet.Range("B11")
			range.Text = "Regular comment"
			range.Comment.Text = "Regular comment"
			range.AutoFitColumns()
			'Regular comment


			range = sheet.Range("B12")
			range.Text = "Rich text comment"
			range.RichText.SetFont(0, 16, font)
			range.AutoFitColumns()
			'Rich text comment
			range.Comment.RichText.Text = "Rich text comment"
			range.Comment.RichText.SetFont(0,4, fontGreen)
			range.Comment.RichText.SetFont(5,9, fontBlue)

			Dim result As String = "WriteComment_result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2007)
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
