Imports Spire.Xls

Namespace CSVToDataTable

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the CSV file from the specified path with the specified delimiter (",")
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CSVSample.csv", ",")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Export the data from the worksheet to a DataTable
            Dim table As DataTable = worksheet.ExportDataTable()

            'Show in data grid
            Me.dataGridView1.DataSource = table
        End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
