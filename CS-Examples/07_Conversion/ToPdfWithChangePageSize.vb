Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO
Imports System.Net

Namespace ToPdfWithChangePageSize
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

            ' Iterate through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Set the paper size of the current worksheet to A3
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
            Next sheet

            ' Specify the filename for the resulting PDF file
            Dim result As String = "result.pdf"

            ' Save the workbook as a PDF file with the specified filename and format
            workbook.SaveToFile(result, FileFormat.PDF)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
