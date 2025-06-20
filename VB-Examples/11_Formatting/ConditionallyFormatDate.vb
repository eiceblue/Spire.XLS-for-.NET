Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet.ConditionalFormatting

Namespace ConditionallyFormatDate
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Load a workbook from a file named "Template_Xls_6.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add conditional formats to the worksheet for the allocated range
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            xcfs.AddRange(sheet.AllocatedRange)

            ' Add a time period condition (last 7 days) to the conditional format
            Dim conditionalFormat As IConditionalFormat = xcfs.AddTimePeriodCondition(TimePeriodType.Last7Days)
            conditionalFormat.BackColor = Color.Orange

            ' Save the workbook to a file named "Result-ConditionallyFormatDate.xlsx" in Excel 2013 format
            Dim result As String = "Result-ConditionallyFormatDate.xlsx"
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
