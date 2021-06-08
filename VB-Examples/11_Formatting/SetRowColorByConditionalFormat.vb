Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace SetRowColorByConditionalFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Select the range that you want to format.
			Dim dataRange As CellRange = sheet.AllocatedRange

			'Set conditional formatting.
			Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs.AddRange(dataRange)
			Dim format1 As IConditionalFormat = xcfs.AddCondition()
			'Determines the cells to format.
			format1.FirstFormula = "=MOD(ROW(),2)=0"
			'Set conditional formatting type
			format1.FormatType = ConditionalFormatType.Formula
			'Set the color.
			format1.BackColor = Color.LightSeaGreen

			'Set the backcolor of the odd rows as Yellow.
			Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs1.AddRange(dataRange)
			Dim format2 As IConditionalFormat = xcfs.AddCondition()
			format2.FirstFormula = "=MOD(ROW(),2)=1"
			format2.FormatType = ConditionalFormatType.Formula
			format2.BackColor = Color.Yellow

			Dim result As String = "Result-SetRowColorWithConditionalFormatting.xlsx"

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
