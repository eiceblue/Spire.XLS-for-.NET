Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Drawing.Imaging
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet

Namespace ToHtmlStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create an instance of HTMLOptions class to specify HTML conversion settings.
            Dim options As New HTMLOptions()
            options.ImageEmbedded = True

            ' Specify the output file name for the HTML file.
            Dim outputFile As String = "Output.html"

            ' Create a FileStream object in FileMode.Create to write the HTML content to a file.
            Dim fileStream As New FileStream(outputFile, FileMode.Create)

            ' Save the worksheet to an HTML file using the specified file stream and options.
            sheet.SaveToHtml(fileStream, options)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
