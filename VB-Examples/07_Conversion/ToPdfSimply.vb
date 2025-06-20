Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core

Namespace ToPdfSimply
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPDF.xlsx")

            ' Specify the file name
            Dim result As String = "result.pdf"

            ' Save the workbook as a PDF file.
            workbook.SaveToFile(result, FileFormat.PDF)

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
