Imports Spire.Xls

Namespace CSVToDataTable

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CSVSample.csv", ",")

			'Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			'Export to datatable
			Dim t As DataTable = worksheet.ExportDataTable()
			'Show in data grid
			Me.dataGridView1.DataSource = t
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
