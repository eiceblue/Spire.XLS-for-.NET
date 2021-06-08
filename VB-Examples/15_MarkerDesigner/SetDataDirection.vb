Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SetDataDirection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner2.xlsx")

			' Create a DataTable
			Dim dt As New DataTable("data")

			'Define a field in it
			dt.Columns.Add(New DataColumn("value", GetType(String)))

			' Add three rows to it
			Dim drName1 As DataRow = dt.NewRow()
			Dim drName2 As DataRow = dt.NewRow()
			Dim drName3 As DataRow = dt.NewRow()

			drName1("value") = "Text1"
			drName2("value") = "Text2"
			drName3("value") = "Text3"


			dt.Rows.Add(drName1)
			dt.Rows.Add(drName2)
			dt.Rows.Add(drName3)

			'Fill DataTable
			workbook.MarkerDesigner.AddDataTable("data", dt)
			workbook.MarkerDesigner.Apply()

			'Save the document
			Dim output As String = "SetDataDirection_result.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'View the document
			FileViewer(output)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
