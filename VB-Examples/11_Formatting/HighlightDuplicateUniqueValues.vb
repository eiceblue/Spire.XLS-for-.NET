Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace HighlightDuplicateUniqueValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new conditional formats collection for the first set of conditions
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Set the range of cells (C2 to C10) to apply the conditional formatting to
            xcfs.AddRange(sheet.Range("C2:C10"))

            ' Add a condition for duplicate values and obtain the corresponding conditional format object
            Dim format1 As IConditionalFormat = xcfs.AddCondition()

            ' Set the format type to highlight duplicate values
            format1.FormatType = ConditionalFormatType.DuplicateValues

            ' Set the background color for cells that satisfy the condition to Indian Red
            format1.BackColor = Color.IndianRed


            ' Create another conditional formats collection for the second set of conditions
            Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Set the range of cells (C2 to C10) for the second set of conditions
            xcfs1.AddRange(sheet.Range("C2:C10"))

            ' Add a condition for unique values and obtain the corresponding conditional format object
            Dim format2 As IConditionalFormat = xcfs.AddCondition()

            ' Set the format type to highlight unique values
            format2.FormatType = ConditionalFormatType.UniqueValues

            ' Set the background color for cells that satisfy the condition to Yellow
            format2.BackColor = Color.Yellow

            ' Specify the file name to save the modified workbook
            Dim result As String = "Result-HighlightDuplicateAndUniqueValues.xlsx"

            ' Save the modified workbook to the specified file path, using Excel 2013 format
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
