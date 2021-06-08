Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace MarkerDesigner
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner.xls")
			Dim dt As DataTable = CType(dataGrid1.DataSource, DataTable)

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Fill parameter
			workbook.MarkerDesigner.AddParameter("Variable1",1234.5678)
			'Fill DataTable
			workbook.MarkerDesigner.AddDataTable("Country",dt)
			workbook.MarkerDesigner.Apply()
			'AutoFit
			sheet.AllocatedRange.AutoFitRows()
			sheet.AllocatedRange.AutoFitColumns()


			Dim result As String = "Output_MarkerDesigner.xlsx"

			workbook.SaveToFile(result, ExcelVersion.Version2010)
			ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner-DataSample.xls")
			'Initailize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Me.dataGrid1.DataSource = sheet.ExportDataTable()
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub


	End Class
End Namespace
