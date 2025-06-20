Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ToCSVWithFilteredValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AutofilterSample.xlsx")

            ' Save the first worksheet of the workbook to a CSV file.
            ' Use semicolon as the delimiter and do not retain hidden data.
            workbook.Worksheets(0).SaveToFile("ToCSVWithFilteredValue.csv", ";", False)

            ' Save the worksheet to a stream with the specified parameters.
            ' Parameters: Stream object, separator (string), and retainHiddenData (boolean).
            ' Note: This line is commented out in the provided code.
            'worksheet.SaveToStream(Stream stream, string separator, bool retainHiddenData);

            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer("ToCSVWithFilteredValue.csv")
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
