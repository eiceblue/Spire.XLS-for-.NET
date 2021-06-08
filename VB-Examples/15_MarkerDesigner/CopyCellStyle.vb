Imports Spire.Xls

Namespace CopyCellStyle

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner1.xlsx")

			' Create Students DataTable
			Dim dt As New DataTable("data")

			' Define a field in it
			dt.Columns.Add(New DataColumn("name", GetType(String)))
			dt.Columns.Add(New DataColumn("age", GetType(Integer)))

			' Add three rows to it
			Dim drName1 As DataRow = dt.NewRow()
			Dim drName2 As DataRow = dt.NewRow()
			Dim drName3 As DataRow = dt.NewRow()

			drName1("name") = "John"
			drName1("age") = 15
			drName2("name") = "Jess"
			drName2("age") = 22
			drName3("name") = "Alan"
			drName3("age") = 36

			dt.Rows.Add(drName1)
			dt.Rows.Add(drName2)
			dt.Rows.Add(drName3)

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Fill DataTable
			workbook.MarkerDesigner.AddDataTable("data", dt)
			workbook.MarkerDesigner.Apply()

			'Save the document
			Dim output As String = "CopyCellStyle.xlsx"
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
