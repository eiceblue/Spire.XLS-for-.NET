Imports Spire.Xls

Namespace AddCommentWithPicture

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Set the text value of cell C6 to "E-iceblue".
            sheet.Range("C6").Text = "E-iceblue"

            'Add a comment to cell C6.
            Dim comment As ExcelComment = sheet.Range("C6").AddComment()

            'Load the image file "Logo.png" into an Image object.
            Dim image As Image = image.FromFile("Logo.png")

            'Set the custom picture fill for the comment using the image and its associated filename.
            comment.Fill.CustomPicture(image, "logo.png")

            'Set the height and width of the comment to match the dimensions of the image.
            comment.Height = image.Height
            comment.Width = image.Width

            'Set the visibility of the comment to true.
            comment.Visible = True

            'Specify the filename to save the modified workbook.
            Dim output As String = "AddCommentWithPicture.xlsx"

            'Save the workbook to the specified file using Excel 2010 format.
            workbook.SaveToFile(output, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
