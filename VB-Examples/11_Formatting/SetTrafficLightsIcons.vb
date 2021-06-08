Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace SetTrafficLightsIcons
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Add a worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add some data to the Excel sheet cell range and set the format for them.
			sheet.Range("A1").Text = "Traffic Lights"
			sheet.Range("A2").NumberValue = 0.95
			sheet.Range("A2").NumberFormat = "0%"
			sheet.Range("A3").NumberValue = 0.5
			sheet.Range("A3").NumberFormat = "0%"
			sheet.Range("A4").NumberValue = 0.1
			sheet.Range("A4").NumberFormat = "0%"
			sheet.Range("A5").NumberValue = 0.9
			sheet.Range("A5").NumberFormat = "0%"
			sheet.Range("A6").NumberValue = 0.7
			sheet.Range("A6").NumberFormat = "0%"
			sheet.Range("A7").NumberValue = 0.6
			sheet.Range("A7").NumberFormat = "0%"

			'Set the height of row and width of column for Excel cell range.
			sheet.AllocatedRange.RowHeight = 20
			sheet.AllocatedRange.ColumnWidth = 25

			'Add a conditional formatting.
			Dim conditional As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			conditional.AddRange(sheet.AllocatedRange)
			Dim format1 As IConditionalFormat = conditional.AddCondition()

			'Add a conditional formatting of cell range and set its type to CellValue.
			format1.FormatType = ConditionalFormatType.CellValue
			format1.FirstFormula = "300"
			format1.Operator = ComparisonOperatorType.Less
			format1.FontColor = Color.Black
			format1.BackColor = Color.LightSkyBlue

			'Add a conditional formatting of cell range and set its type to IconSet.
			conditional.AddRange(sheet.AllocatedRange)
			Dim format As IConditionalFormat = conditional.AddCondition()
			format.FormatType = ConditionalFormatType.IconSet
			format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1

			Dim result As String = "Result-SetTrafficLightsIcons.xlsx"

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
