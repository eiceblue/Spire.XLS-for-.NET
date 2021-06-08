Imports Spire.Xls

Namespace Subtotal
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Subtotal.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Select data range
			Dim range As CellRange = sheet.Range("A1:B18")
			'Subtotal selected data
			sheet.Subtotal(range, 0, New Integer() {1}, SubtotalTypes.Sum, True, False, True)

			'Save to file
			Dim result As String = "Subtotal_Out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
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
