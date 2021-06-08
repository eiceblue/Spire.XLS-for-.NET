Imports Spire.Xls

Namespace SetPageBreak

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set Excel Page Break Horizontally
			sheet.HPageBreaks.Add(sheet.Range("A8"))
			sheet.HPageBreaks.Add(sheet.Range("A14"))

			'Set Excel Page Break Vertically
			'sheet.VPageBreaks.Add(sheet.Range["B1"]);
			'sheet.VPageBreaks.Add(sheet.Range["C1"]);

			'Set view mode to Preview mode
			workbook.Worksheets(0).ViewMode = ViewMode.Preview

			'Save the document
			Dim output As String = "SetPageBreak.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

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
