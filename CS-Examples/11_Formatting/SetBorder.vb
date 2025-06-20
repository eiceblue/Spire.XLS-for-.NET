Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetBorder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SetBorder.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Define a CellRange object that includes all cells in the worksheet
            Dim cr As CellRange = sheet.Range(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)

            ' Set the border style of the CellRange to Double line
            cr.Borders.LineStyle = LineStyleType.Double

            ' Set the diagonal down border style of the CellRange to None
            cr.Borders(BordersLineType.DiagonalDown).LineStyle = LineStyleType.None

            ' Set the diagonal up border style of the CellRange to None
            cr.Borders(BordersLineType.DiagonalUp).LineStyle = LineStyleType.None

            ' Set the border color of the CellRange to CadetBlue
            cr.Borders.Color = Color.CadetBlue

            ' Specify the output file name for the modified workbook
            Dim result As String = "SetBorder_result.xlsx"

            ' Save the modified workbook to the specified file path with Excel version 2010
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
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
