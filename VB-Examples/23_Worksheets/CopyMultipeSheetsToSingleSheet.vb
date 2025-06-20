Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace CopyMultipeSheetsToSingleSheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load a workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_11.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            ' Iterate through each worksheet except the first one
            For i As Integer = 1 To workbook.Worksheets.Count - 1
                ' Get the current worksheet
                Dim sheet2 As Worksheet = workbook.Worksheets(i)

                ' Copy the cells from the current worksheet to the first worksheet
                sheet2.Copy(CType(sheet2.MaxDisplayRange, CellRange), sheet1, sheet1.LastRow + 1, sheet2.FirstColumn, True)
            Next i

            ' Specify the output file name
            Dim fileName As String = "CopyMultipeSheetsToSingleSheet_result.xlsx"

            ' Save the workbook to the specified file path in Excel 2016 format
            workbook.SaveToFile(fileName, ExcelVersion.Version2016)

            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer(fileName)
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
