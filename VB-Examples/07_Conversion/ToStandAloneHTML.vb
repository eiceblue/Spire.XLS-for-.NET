Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace ToStandAloneHTML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToCSV.xlsx")

            ' Set the IsStandAloneHtmlFile property of HTMLOptions.Default to True
            HTMLOptions.Default.IsStandAloneHtmlFile = True

            ' Specify the output filename for the HTML file
            Dim outputFile As String = "Output.html"

            ' Create a FileStream object to write the HTML file
            Dim fileStream As New FileStream(outputFile, FileMode.Create)

            ' Save the workbook as an HTML file to the FileStream using the specified format (HTML)
            workbook.SaveToStream(fileStream, FileFormat.HTML)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'view the document
            ExcelDocViewer(outputFile)
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
