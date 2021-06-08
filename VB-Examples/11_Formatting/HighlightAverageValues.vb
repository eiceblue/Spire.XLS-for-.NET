Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace HighlightAverageValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add conditional format.
			Dim format1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			'Set the cell range to apply the formatting.
			format1.AddRange(sheet.Range("E2:E10"))
			'Add below average condition.
			Dim cf1 As IConditionalFormat = format1.AddAverageCondition(AverageType.Below)
			'Highlight cells below average values.
			cf1.BackColor = Color.SkyBlue

			'Add conditional format.
			Dim format2 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			'Set the cell range to apply the formatting.
			format2.AddRange(sheet.Range("E2:E10"))
			'Add above average condition.
			Dim cf2 As IConditionalFormat = format1.AddAverageCondition(AverageType.Above)
			'Highlight cells above average values.
			cf2.BackColor = Color.Orange

			Dim result As String = "Result-HighlightBelowAndAboveAverageValues.xlsx"

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
