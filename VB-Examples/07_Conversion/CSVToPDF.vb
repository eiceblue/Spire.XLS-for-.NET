Imports Spire.Xls

Namespace CSVToPDF

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CSVSample.csv",",", 1, 1)

			'Set the SheetFitToPage property as true
			workbook.ConverterSetting.SheetFitToPage = True

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Autofit a column if the characters in the column exceed column width
			For i As Integer = 1 To sheet.Columns.Length - 1
				sheet.AutoFitColumn(i)
			Next i

			'Save to PDF document
			Dim output As String = "CSVToPDF.pdf"
			workbook.SaveToFile(output, FileFormat.PDF)

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
