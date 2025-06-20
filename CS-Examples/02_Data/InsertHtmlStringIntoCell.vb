Imports Spire.Xls

Namespace InsertHtmlStringIntoCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first sheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Define an HTML code string
            Dim htmlCode As String = "<div>first line<br>second line<br>third line</div>"

            ' Access the cell range A1 in the worksheet
            Dim range As CellRange = sheet("A1")

            ' Set the HTML string as the value of the cell range
            range.HtmlString = htmlCode

            ' Specify the filename for saving the workbook
            Dim result As String = "InsertHtmlStringIntoCell.xlsx"

            ' Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
