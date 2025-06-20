Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetTabColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SetTabColor.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Set the tab color of the first worksheet to red
            worksheet.TabColor = Color.Red

            ' Get the second worksheet from the workbook
            worksheet = workbook.Worksheets(1)

            ' Set the tab color of the second worksheet to green
            worksheet.TabColor = Color.Green

            ' Get the third worksheet from the workbook
            worksheet = workbook.Worksheets(2)

            ' Set the tab color of the third worksheet to light blue
            worksheet.TabColor = Color.LightBlue

            ' Save the modified workbook to a file in Excel 2010 format
            workbook.SaveToFile("SetTabColor_result.xlsx", ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("SetTabColor_result.xlsx")
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
