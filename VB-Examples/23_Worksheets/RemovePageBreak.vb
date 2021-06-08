Imports Spire.Xls

Namespace RemovePageBreak

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\PageBreak.xlsx")

			'Get the first worksheet from the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Clear all the vertical page breaks
			sheet.VPageBreaks.Clear()

			'Remove the firt horizontal Page Break
			sheet.HPageBreaks.RemoveAt(0)

			'Set the ViewMode as Preview to see how the page breaks work
			sheet.ViewMode = ViewMode.Preview

			'Save the document
			Dim output As String = "RemovePageBreak.xlsx"
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
