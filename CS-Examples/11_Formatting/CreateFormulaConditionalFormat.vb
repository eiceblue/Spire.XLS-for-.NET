Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace CreateFormulaConditionalFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get a range of cells in the first column of the worksheet
            Dim range As CellRange = sheet.Columns(0)

            ' Add conditional formatting to the worksheet
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            xcfs.AddRange(range)

            ' Add a condition to the conditional formatting
            Dim conditional As IConditionalFormat = xcfs.AddCondition()
            conditional.FormatType = ConditionalFormatType.Formula
            conditional.FirstFormula = "=($A1<$B1)"
            conditional.BackKnownColor = ExcelColors.Yellow

            ' Save the modified workbook to a new file
            Dim result As String = "Result-CreateFormulaToApplyConditionalFormatting.xlsx"
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
