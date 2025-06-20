Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace HideMajorGridlines
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

            ' Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load the Excel file into the workbook.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampeB_4.xlsx")

            ' Get the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart in the worksheet.
            Dim chart As Chart = sheet.Charts(0)

            ' Disable major grid lines on the primary value axis of the chart.
            chart.PrimaryValueAxis.HasMajorGridLines = False

            ' Specify the result file name as "result.xlsx".
            Dim result As String = "result.xlsx"

            ' Save the modified workbook to the result file using Excel 2010 format.
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
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
