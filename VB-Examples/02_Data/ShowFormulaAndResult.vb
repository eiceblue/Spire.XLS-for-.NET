Imports Spire.Xls

Namespace ShowFormulaAndResult

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		'Formula
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FormulasSample.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Show formula
			Dim dt As DataTable = sheet.ExportDataTable(sheet.AllocatedRange, False, False)
			'Show in DataGridView
			Me.dataGridView1.DataSource = dt
		End Sub

		'Result
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FormulasSample.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Show result
			Dim dt As DataTable = sheet.ExportDataTable(sheet.AllocatedRange, False, True)
			'//Show in DataGridView
			Me.dataGridView1.DataSource = dt
		End Sub

	End Class
End Namespace
