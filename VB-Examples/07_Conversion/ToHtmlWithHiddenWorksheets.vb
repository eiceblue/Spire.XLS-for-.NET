Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ToHtmlWithHiddenWorksheets
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim book As New Workbook()

            ' Load the Excel file from the specified path.
            book.LoadFromFile("..\..\..\..\..\..\Data\ToHtmlWithHiddenWorksheets.xlsx")

            ' Specify the output HTML file name.
            Dim result As String = "result.html"

            ' Save the workbook to an HTML file.
            ' Parameter: 
            ' - False: Save to HTML with hidden worksheets included.
            ' - True: Save to HTML without hidden worksheets.
            book.SaveToHtml(result, False)
            ' Release the resources used by the workbook
            book.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
