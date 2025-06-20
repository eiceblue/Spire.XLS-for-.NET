Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace DataExport
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DataExport.xlsx")

            ' Initialize a Worksheet object by getting the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the data source of a DataGrid (assuming "dataGrid1" is a DataGrid control) to the exported DataTable from the worksheet
            Me.dataGrid1.DataSource = sheet.ExportDataTable()

            ' Release the resources used by the workbook
            workbook.Dispose()

        End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub



	End Class
End Namespace
