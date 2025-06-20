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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new conditional format object
            Dim format1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Set the range of cells to apply the conditional formatting to
            format1.AddRange(sheet.Range("E2:E10"))

            ' Create a new conditional format for below average values
            Dim cf1 As IConditionalFormat = format1.AddAverageCondition(AverageType.Below)

            ' Set the background color of the cells that satisfy the condition
            cf1.BackColor = Color.SkyBlue

            ' Create another conditional format object
            Dim format2 As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Set the range of cells for the second conditional format
            format2.AddRange(sheet.Range("E2:E10"))

            ' Create a new conditional format for above average values
            Dim cf2 As IConditionalFormat = format1.AddAverageCondition(AverageType.Above)

            ' Set the background color of the cells that satisfy the second condition
            cf2.BackColor = Color.Orange

            ' Specify the file name to save the modified workbook
            Dim result As String = "Result-HighlightBelowAndAboveAverageValues.xlsx"

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
