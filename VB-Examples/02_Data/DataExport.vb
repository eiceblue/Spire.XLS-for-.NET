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
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\DataExport.xlsx")
			'Initailize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Me.dataGrid1.DataSource = sheet.ExportDataTable()
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub



	End Class
End Namespace
