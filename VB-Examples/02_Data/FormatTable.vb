Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace FormatTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FormatTable.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Add Default Style to the table
			sheet.ListObjects(0).BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9
			'Show Total
			sheet.ListObjects(0).DisplayTotalRow = True
			'Set calculation type
			sheet.ListObjects(0).Columns(0).TotalsRowLabel = "Total"
			sheet.ListObjects(0).Columns(1).TotalsCalculation = ExcelTotalsCalculation.None
			sheet.ListObjects(0).Columns(2).TotalsCalculation = ExcelTotalsCalculation.None
			sheet.ListObjects(0).Columns(3).TotalsCalculation = ExcelTotalsCalculation.Sum
			sheet.ListObjects(0).Columns(4).TotalsCalculation = ExcelTotalsCalculation.Sum

			sheet.ListObjects(0).ShowTableStyleRowStripes = True

			sheet.ListObjects(0).ShowTableStyleColumnStripes = True
			workbook.SaveToFile("Sample.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("Sample.xlsx")
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
