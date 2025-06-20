Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace ApplyConditionalFormatting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set numeric values for cells A1 to C4 in the worksheet
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

            ' Set the row height and column width for the allocated range in the worksheet
            sheet.AllocatedRange.RowHeight = 15
            sheet.AllocatedRange.ColumnWidth = 17

            ' Add the first conditional formatting rule to the worksheet based on cell values
            Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            xcfs1.AddRange(sheet.AllocatedRange)
            Dim format1 As IConditionalFormat = xcfs1.AddCondition()
            format1.FormatType = ConditionalFormatType.CellValue
            format1.FirstFormula = "800"
            format1.Operator = ComparisonOperatorType.Greater
            format1.FontColor = Color.Red
            format1.BackColor = Color.LightSalmon

            ' Add the second conditional formatting rule to the worksheet based on cell values
            Dim xcfs2 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            xcfs2.AddRange(sheet.AllocatedRange)
            Dim format2 As IConditionalFormat = xcfs2.AddCondition()
            format2.FormatType = ConditionalFormatType.CellValue
            format2.FirstFormula = "300"
            format2.Operator = ComparisonOperatorType.Less
            format2.FontColor = Color.Green
            format2.BackColor = Color.LightBlue

            ' Specify the output file path for saving the modified workbook
            Dim result As String = "Result-ApplyConditionalFormattingToDataRange.xlsx"

            ' Save the workbook to the specified output file path in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
