Imports Spire.Xls
Imports Spire.Xls.Collections

Namespace RemoveComment

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CommentSample.xlsx")

            ' Retrieve the CommentsCollection from the first worksheet of the workbook
            Dim comments As CommentsCollection = workbook.Worksheets(0).Comments

            ' Modify the text of the first comment
            comments(0).Text = "This comment has been changed."

            ' Remove the second comment from the collection
            comments(1).Remove()

            ' Specify the output file name for saving the modified workbook
            Dim output As String = "RemoveAndChangeComment.xlsx"

            ' Save the workbook to the specified file path using Excel 2013 format
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
