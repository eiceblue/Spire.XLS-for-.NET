Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace MakeCellActive
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(1)

            ' Activate the worksheet
            sheet.Activate()

            ' Set the active cell to cell B2
            sheet.SetActiveCell(sheet.Range("B2"))

            ' Set the first visible column to column 1
            sheet.FirstVisibleColumn = 1

            ' Set the first visible row to row 1
            sheet.FirstVisibleRow = 1

            ' Specify the filename for the resulting Excel file
            Dim result As String = "MakeCellActive_result.xlsx"

            ' Save the workbook to the specified filename in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
