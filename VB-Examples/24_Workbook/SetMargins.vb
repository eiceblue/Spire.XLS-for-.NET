Imports Spire.Xls

Namespace SetMargins

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

			'Set margins for top, bottom, left and right, here the unit of measure is Inch
			sheet.PageSetup.TopMargin = 0.3
			sheet.PageSetup.BottomMargin = 1
			sheet.PageSetup.LeftMargin = 0.2
			sheet.PageSetup.RightMargin = 1
			'Set the header margin and footer margin
			sheet.PageSetup.HeaderMarginInch = 0.1
			sheet.PageSetup.FooterMarginInch = 0.5

			'Save the document
			Dim output As String = "SetMargins.xlsx"
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
