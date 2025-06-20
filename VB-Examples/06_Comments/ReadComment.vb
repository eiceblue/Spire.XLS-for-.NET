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
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadComment.xls")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Get the text of the comment in cell A1.
            textBox1.Text = sheet.Range("A1").Comment.Text

            'Get the RTF text of the comment in cell A2.
            richTextBox1.Rtf = sheet.Range("A2").Comment.RichText.RtfText
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
