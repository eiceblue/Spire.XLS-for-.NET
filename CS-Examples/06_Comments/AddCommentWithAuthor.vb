Imports Spire.Xls

Namespace AddCommentWithAuthor

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Get the CellRange object for cell C1.
            Dim range As CellRange = sheet.Range("C1")

            'Specify the author and text for the comment.
            Dim author As String = "E-iceblue"
            Dim text As String = "This is demo to show how to add a comment with editable Author property."

            'Add a comment to the range and customize its properties.
            Dim comment As ExcelComment = range.AddComment()
            comment.Width = 200
            comment.Visible = True
            comment.Text = If(String.IsNullOrEmpty(author), text, author & ":" & vbLf & text)

            'Create a new font and apply it to the author portion of the comment's rich text.
            Dim font As ExcelFont = range.Worksheet.Workbook.CreateFont()
            font.FontName = "Tahoma"
            font.KnownColor = ExcelColors.Black
            font.IsBold = True
            comment.RichText.SetFont(0, author.Length, font)

            'Specify the filename to save the modified workbook.
            Dim output As String = "AddCommentWithAuthor.xlsx"

            'Save the workbook to the specified file using Excel 2013 format.
            workbook.SaveToFile(output, ExcelVersion.Version2013)
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
