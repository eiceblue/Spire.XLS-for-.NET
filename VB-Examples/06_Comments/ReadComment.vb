Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace ReadComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()

					  workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadComment.xls")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			textBox1.Text = sheet.Range("A1").Comment.Text
			richTextBox1.Rtf = sheet.Range("A2").Comment.RichText.RtfText
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
