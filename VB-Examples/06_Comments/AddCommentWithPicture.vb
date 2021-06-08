Imports Spire.Xls

Namespace AddCommentWithPicture

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("C6").Text = "E-iceblue"
			'Add the comment
			Dim comment As ExcelComment = sheet.Range("C6").AddComment()
			'Load the image file
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Logo.png")
			'Fill the comment with a customized background picture
			comment.Fill.CustomPicture(image, "logo.png")
			'Set the height and width of comment
			comment.Height = image.Height
			comment.Width = image.Width
			comment.Visible = True

			'Save the document
			Dim output As String = "AddCommentWithPicture.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2010)

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
