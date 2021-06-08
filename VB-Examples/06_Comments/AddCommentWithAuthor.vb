Imports Spire.Xls

Namespace AddCommentWithAuthor

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the range that will add comment
			Dim range As CellRange = sheet.Range("C1")

			'Set the author and comment content
			Dim author As String = "E-iceblue"
			Dim text As String = "This is demo to show how to add a comment with editable Author property."

			'Add comment to the range and set properties
			Dim comment As ExcelComment = range.AddComment()
			comment.Width = 200
			comment.Visible = True
			comment.Text = If(String.IsNullOrEmpty(author), text, author & ":" & vbLf & text)

			'Set the font of the author
			Dim font As ExcelFont = range.Worksheet.Workbook.CreateFont()
			font.FontName = "Tahoma"
			font.KnownColor = ExcelColors.Black
			font.IsBold = True
			comment.RichText.SetFont(0, author.Length, font)

			'Save the document
			Dim output As String = "AddCommentWithAuthor.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
