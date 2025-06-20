Imports Spire.Xls
Imports Spire.Xls.Pdf
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ToPDFA1B
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPDF_A1BExample.xlsx")

            ' Set the PDF conformance level to A1B for the Workbook's ConverterSetting object
            workbook.ConverterSetting.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B

            ' Save the Workbook as a PDF file with the specified filename and format
            workbook.SaveToFile("ToPDFA1B_result.pdf", FileFormat.PDF)
            ' Release the resources used by the workbook
            workbook.Dispose()

            FileViewer("ToPDFA1B_result.pdf")
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
