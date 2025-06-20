Imports Spire.Xls
Imports System.IO

Namespace ObtainActiveSelectionRange
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Load an existing workbook from a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ObtainActiveSelectionRange.xlsx")

			' Get the first sheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			Dim information As String = Nothing

			' Get the information of the active selection range
			For Each range As CellRange In worksheet.ActiveSelectionRange
				information &= "RangeAddressLocal:" & range.RangeAddressLocal & vbCrLf
				information &= "ColumnCount:" & range.ColumnCount & vbCrLf
				information &= "ColumnWidth:" & range.ColumnWidth & vbCrLf
				information &= "Column:" & range.Column & vbCrLf
				information &= "RowCount:" & range.RowCount & vbCrLf
				information &= "RowHeight:" & range.RowHeight & vbCrLf
				information &= "Row:" & range.Row & vbCrLf
			Next range

			' Specify the output file name for the result
			Dim result As String = "ObtainActiveSelectionRange_result.txt"

			' Write the content of the information to the result file
			File.WriteAllText(result, information)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
