Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core


Namespace FormatNamedRangeCells
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AllNamedRanges.xlsx")

            ' Get the first named range from the workbook
            Dim NamedRange As INamedRange = workbook.NameRanges(0)

            ' Get the range referred by the named range
            Dim range As IXLSRange = NamedRange.RefersToRange

            ' Set the color of the range to yellow
            range.Style.Color = Color.Yellow

            ' Make the font bold for the range
            range.Style.Font.IsBold = True

            ' Define the output file name as "result.xlsx"
            Dim result As String = "result.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2010 format
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
