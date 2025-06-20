Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace FitWidthWhenConvertToPDF
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

            ' Iterate through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets
                ' Set the number of pages tall to 0 for automatic scaling
                sheet.PageSetup.FitToPagesTall = 0

                ' Set the number of pages wide to 1 for fitting the content within one page width
                sheet.PageSetup.FitToPagesWide = 1
            Next sheet

            ' Specify the output file name for saving as PDF
            Dim result As String = "result.pdf"

            ' Save the workbook to a PDF file
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
