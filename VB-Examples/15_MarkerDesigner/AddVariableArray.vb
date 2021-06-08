Imports Spire.Xls

Namespace AddVariableArray

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set marker designer field in cell A1
			sheet.Range("A1").Value = "&=Array"

			'Fill Array
			workbook.MarkerDesigner.AddArray("Array", New String() { "Spire.Xls", "Spire.Doc", "Spire.PDF", "Spire.Presentation", "Spire.Email" })
			workbook.MarkerDesigner.Apply()
			workbook.CalculateAllValue()

			'AutoFit
			sheet.AllocatedRange.AutoFitRows()
			sheet.AllocatedRange.AutoFitColumns()

			'Save the document
			Dim output As String = "AddVariableArray.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the file
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
