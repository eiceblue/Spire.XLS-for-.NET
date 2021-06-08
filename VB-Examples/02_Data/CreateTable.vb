Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace CreateTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Add a new List Object to the worksheet
			sheet.ListObjects.Create("table", sheet.Range(1, 1, 19, 5))
			' Add Default Style to the table    
			sheet.ListObjects(0).BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9

			'Save to file
			Dim result As String = "CreateTable_out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
