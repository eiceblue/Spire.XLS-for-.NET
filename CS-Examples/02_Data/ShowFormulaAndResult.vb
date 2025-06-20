Imports Spire.Xls

Namespace ShowFormulaAndResult

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		'Formula
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Creates a new Excel workbook.
            Dim workbook As New Workbook()

            ' Loads an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\FormulasSample.xlsx")
            ' Retrieves the first worksheet in the workbook (index starts at 0).
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Exports the data from the worksheet to a DataTable object.
            'The AllocatedRange represents the range of cells that contain data in the worksheet.
            'The first False parameters are used to exclude column headers from the exported data.
            'The second False parameters are used to show formulas from the exported data.
            Dim dt As DataTable = sheet.ExportDataTable(sheet.AllocatedRange, False, False)

            'Assigns the DataTable as the data source for a DataGridView control, displaying the data in tabular form.
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
