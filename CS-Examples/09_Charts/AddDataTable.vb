Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace AddDataTable
	''' <summary>
	''' Summary description for Form1.
	''' </summary>
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "AddDataTable.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AddDataTable.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart (index 0) on the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Enable the display of a data table for the chart
            chart.HasDataTable = True

            ' Save the modified workbook to a new Excel file named "Output.xlsx" with the Excel 2010 format
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Output.xlsx")
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
