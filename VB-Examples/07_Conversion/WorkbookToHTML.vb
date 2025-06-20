Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace WorkbookToHTML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorkbookToHTML.xlsx")

            ' Save the workbook to an HTML file.
            workbook.SaveToHtml("result.html")
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer("result.html")
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
