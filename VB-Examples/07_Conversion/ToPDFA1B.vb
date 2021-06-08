Imports Spire.Xls
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
			'Create a workbook
			Dim workbook As New Workbook()

			'Load an excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPDF_A1BExample.xlsx")

			'Convert excel to PDFA/1-B
			workbook.ConverterSetting.PdfConformanceLevel = Spire.Pdf.PdfConformanceLevel.Pdf_A1B

			'Save the document and launch it
			workbook.SaveToFile("ToPDFA1B_result.pdf", FileFormat.PDF)

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
