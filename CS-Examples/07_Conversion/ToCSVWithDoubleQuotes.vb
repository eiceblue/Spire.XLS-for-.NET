Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ToCSVWithDoubleQuotes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToCSV.xlsx")

            ' Convert the workbook to a CSV file.
            ' When the last parameter is set to true, double quotes will be added around each field. The default value is false.
            workbook.SaveToFile("ToCSVAddQuotation.csv", ",", True)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'view the document
            ExcelDocViewer("ToCSVAddQuotation.csv")
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
