Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace ApplyIconSetsToCellRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Insert data to cell range from A1 to C4.
			sheet.Range("A1").NumberValue = 582
			sheet.Range("A2").NumberValue = 234
			sheet.Range("A3").NumberValue = 314
			sheet.Range("A4").NumberValue = 50
			sheet.Range("B1").NumberValue = 150
			sheet.Range("B2").NumberValue = 894
			sheet.Range("B3").NumberValue = 560
			sheet.Range("B4").NumberValue = 900
			sheet.Range("C1").NumberValue = 134
			sheet.Range("C2").NumberValue = 700
			sheet.Range("C3").NumberValue = 920
			sheet.Range("C4").NumberValue = 450
			sheet.AllocatedRange.RowHeight = 15
			sheet.AllocatedRange.ColumnWidth = 17

			'Add icon sets.
			Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs.AddRange(sheet.AllocatedRange)
			Dim format As IConditionalFormat = xcfs.AddCondition()
			format.FormatType = ConditionalFormatType.IconSet
			format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1

			Dim result As String = "Result-ApplyIconSetsToDataRange.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the MS Excel file.
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
