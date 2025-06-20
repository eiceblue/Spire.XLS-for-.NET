Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace HelloWorld
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()
            ' Access the first sheet of the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)
            ' Set the value of cell A1 in the sheet to "Hello World".
            sheet.Range("A1").Text = "Hello World"
            ' Automatically adjust the width of column A to fit the content.
            sheet.Range("A1").AutoFitColumns()
            ' Specify the file name for the resulting Excel file.
            Dim result As String = "HelloWorld.xlsx"

            ' Save the workbook to a file.
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
