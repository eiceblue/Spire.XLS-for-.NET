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
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Highlight cells that contain a date occurring in the last 7 days.
			Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs.AddRange(sheet.AllocatedRange)
			Dim conditionalFormat As IConditionalFormat = xcfs.AddTimePeriodCondition(TimePeriodType.Last7Days)
			conditionalFormat.BackColor = Color.Orange

			Dim result As String = "Result-ConditionallyFormatDate.xlsx"

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
