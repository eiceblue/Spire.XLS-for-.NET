Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SplitWorksheetIntoPanes

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the index of the first visible column to 2
            sheet.FirstVisibleColumn = 2

            ' Set the index of the first visible row to 5
            sheet.FirstVisibleRow = 5

            ' Set the vertical split position to 4000
            sheet.VerticalSplit = 4000

            ' Set the horizontal split position to 5000
            sheet.HorizontalSplit = 5000

            ' Set the active pane to the first pane
            sheet.ActivePane = 1

            ' Specify the output file name
            Dim output As String = "SplitWorksheetIntoPanes.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
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
