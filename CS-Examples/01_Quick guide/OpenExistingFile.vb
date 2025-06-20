Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace OpenExistingFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()
            ' Load an existing Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz2.xlsx")

            ' Create a new sheet in the workbook and assign it the name "MySheet".
            Dim sheet As Worksheet = workbook.Worksheets.Add("MySheet")

            ' Set the value of cell A1 in the sheet to "Hello World".
            sheet.Range("A1").Text = "Hello World"

            ' Specify the file name for the resulting Excel file.
            Dim result As String = "OpenExistingFile_result.xlsx"

            'Save the workbook to a file.
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
