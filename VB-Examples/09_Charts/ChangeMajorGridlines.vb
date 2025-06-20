Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Imports System.Text
Imports System.IO

Namespace ChangeMajorGridlines
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "SampeB_4.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampeB_4.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Set the color of the major grid lines on the primary value axis to red
            chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.Red

            ' Specify the output file name as "result.xlsx"
            Dim result As String = "result.xlsx"

            ' Save the modified workbook to the specified file path, using the Excel 2010 format
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
