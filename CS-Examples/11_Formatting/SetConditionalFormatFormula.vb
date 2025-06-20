Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Drawing.Imaging
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.Collections

Namespace SetConditionalFormatFormula
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add conditional formats to the worksheet
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Specify the range for conditional formatting
            xcfs.AddRange(sheet.Range("B5"))

            ' Add a new condition for cell value-based formatting
            Dim format As IConditionalFormat = xcfs.AddCondition()
            format.FormatType = ConditionalFormatType.CellValue

            ' Set the condition formula, operator, and background color
            format.FirstFormula = "1000"
            format.Operator = ComparisonOperatorType.Greater
            format.BackColor = Color.Orange

            ' Set the numeric values for cells B1 to B4
            sheet.Range("B1").NumberValue = 40
            sheet.Range("B2").NumberValue = 500
            sheet.Range("B3").NumberValue = 300
            sheet.Range("B4").NumberValue = 400

            ' Set the formula for cell B5 to calculate the sum of B1:B4
            sheet.Range("B5").Formula = "=SUM(B1:B4)"

            ' Provide a description for the purpose of the conditional formatting
            sheet.Range("C5").Text = "If Sum of B1:B4 is greater than 1000, B5 will have orange background."

            ' Save the workbook to an output file in Excel 2013 format
            Dim outputFile As String = "Output.xlsx"
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
