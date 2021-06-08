Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls

Namespace ImportDataFromDataView
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Create an empty worksheet
			workbook.CreateEmptySheets(1)

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create a DataTable object 
			Dim dataTable As New DataTable("Customer")
			dataTable.Columns.Add("No", GetType(Int32))
			dataTable.Columns.Add("Name", GetType(String))
			dataTable.Columns.Add("City", GetType(String))

			'Create rows and add data
			Dim dr As DataRow = dataTable.NewRow()
			dr(0) = 1
			dr(1) = "Tom"
			dr(2) = "New York"
			dataTable.Rows.Add(dr)
			dr = dataTable.NewRow()
			dr(0) = 2
			dr(1) = "Jerry"
			dr(2) = "China"
			dataTable.Rows.Add(dr)
			dr = dataTable.NewRow()
			dr(0) = 3
			dr(1) = "Dive Time"
			dr(2) = "Berkely"
			dataTable.Rows.Add(dr)
			dr = dataTable.NewRow()
			dr(0) = 4
			dr(1) = "Amor Aqua"
			dr(2) = "Florida"
			dataTable.Rows.Add(dr)

			'Import the data view of data table to worksheet
			sheet.InsertDataView(dataTable.DefaultView, True, 1, 1)

			'Save the Excel file
			Dim result As String = "ImportDataFromDataView_output.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(result)
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
