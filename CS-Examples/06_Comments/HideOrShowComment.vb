Imports Spire.Xls

Namespace HideOrShowComment

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CommentSample.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Hide the comment at index 1 in the worksheet.
            sheet.Comments(1).IsVisible = False

            'Show the comment at index 2 in the worksheet.
            sheet.Comments(2).IsVisible = True

            'Specify the filename to save the modified workbook.
            Dim output As String = "HideOrShowComment.xlsx"

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
