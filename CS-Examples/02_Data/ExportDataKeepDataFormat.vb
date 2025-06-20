Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ExportDataKeepDataFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Instantiate a new workbook object.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExportDataKeepDataFormat.xlsx")

            ' Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Instantiate a new ExportTableOptions object to specify export options.
            Dim options As New ExportTableOptions()

            ' Set the option to not keep the data format during export.
            options.KeepDataFormat = False

            'Set the option to rename strategy as "Digit" during export.
            options.RenameStrategy = RenameStrategy.Digit

            ' Export the data from the specified range of the worksheet to a DataTable object.
            Dim table As DataTable = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options)

            ' Set the DataGridView's data source as the exported DataTable.
            Me.dataGridView1.DataSource = table
		End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
