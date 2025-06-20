Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace HtmlToExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Specify the file path of the HTML file to be loaded
            Dim filePath As String = "..\..\..\..\..\..\Data\HtmlToExcel.html"

            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the HTML file into the workbook
            workbook.LoadFromHtml(filePath)

            ' Specify the output file name for saving as Excel file
            Dim result As String = "HtmlToExcel_result.xlsx"

            ' Save the workbook to an Excel file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            ExcelDocViewer(result)
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
