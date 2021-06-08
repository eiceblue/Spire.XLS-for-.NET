Imports Spire.Xls

Namespace SelectedRangeToPDF

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

			'Add a new sheet to workbook
			workbook.Worksheets.Add("newsheet")
			'Copy your area to new sheet.
			workbook.Worksheets(0).Range("A9:E15").Copy(workbook.Worksheets(1).Range("A9:E15"), False, True)
			'Auto fit column width
			workbook.Worksheets(1).Range("A9:E15").AutoFitColumns()

			'Save the document
			Dim output As String = "SelectedRangeToPDF.pdf"
			workbook.Worksheets(1).SaveToPdf(output)

			'Launch the Excel file
			ExcelDocViewer(output)
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
