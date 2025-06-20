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
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text in cell A1
            sheet.Range("A1").Text = "Traffic Lights"

            ' Set the numeric value and number format in cell A2
            sheet.Range("A2").NumberValue = 0.95
            sheet.Range("A2").NumberFormat = "0%"

            ' Set the numeric value and number format in cell A3
            sheet.Range("A3").NumberValue = 0.5
            sheet.Range("A3").NumberFormat = "0%"

            ' Set the numeric value and number format in cell A4
            sheet.Range("A4").NumberValue = 0.1
            sheet.Range("A4").NumberFormat = "0%"

            ' Set the numeric value and number format in cell A5
            sheet.Range("A5").NumberValue = 0.9
            sheet.Range("A5").NumberFormat = "0%"

            ' Set the numeric value and number format in cell A6
            sheet.Range("A6").NumberValue = 0.7
            sheet.Range("A6").NumberFormat = "0%"

            ' Set the numeric value and number format in cell A7
            sheet.Range("A7").NumberValue = 0.6
            sheet.Range("A7").NumberFormat = "0%"

            ' Set the row height and column width for the allocated range
            sheet.AllocatedRange.RowHeight = 20
            sheet.AllocatedRange.ColumnWidth = 25

            ' Add conditional formatting to the sheet
            Dim conditional As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Apply conditional formatting to the allocated range
            conditional.AddRange(sheet.AllocatedRange)

            ' Add a condition for the conditional formatting
            Dim format1 As IConditionalFormat = conditional.AddCondition()

            ' Set the format type to cell value and define the condition
            format1.FormatType = ConditionalFormatType.CellValue
            format1.FirstFormula = "300"
            format1.Operator = ComparisonOperatorType.Less
            format1.FontColor = Color.Black
            format1.BackColor = Color.LightSkyBlue

            ' Apply conditional formatting to the allocated range again
            conditional.AddRange(sheet.AllocatedRange)

            ' Add another condition for the conditional formatting
            Dim format As IConditionalFormat = conditional.AddCondition()

            ' Set the format type to icon set
            format.FormatType = ConditionalFormatType.IconSet
            format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1

            ' Save the workbook to a file
            Dim result As String = "Result-SetTrafficLightsIcons.xlsx"
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
