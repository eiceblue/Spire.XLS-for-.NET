Imports Spire.Xls
Imports Spire.Xls.Collections

Namespace RemoveHyperlinks

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\HyperlinksSample1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the collection of all hyperlinks in the worksheet
			Dim links As HyperLinksCollection = sheet.HyperLinks

			'Remove all link content
			sheet.Range("B1").ClearAll()
			sheet.Range("B2").ClearAll()
			sheet.Range("B3").ClearAll()

			'Remove hyperlink and keep link text
			sheet.HyperLinks.RemoveAt(0)
			sheet.HyperLinks.RemoveAt(0)
			sheet.HyperLinks.RemoveAt(0)

			'Save the document
			Dim output As String = "RemoveHyperlinks.xlsx"
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
