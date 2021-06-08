Imports Spire.Xls
Imports Spire.Xls.Collections

Namespace RemoveComment

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CommentSample.xlsx")

			'Get all comments of the first sheet
			Dim comments As CommentsCollection = workbook.Worksheets(0).Comments
			'Change the content of the first comment
			comments(0).Text = "This comment has been changed."
			'Remove the second comment
			comments(1).Remove()

			'Save the document
			Dim output As String = "RemoveAndChangeComment.xlsx"
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
