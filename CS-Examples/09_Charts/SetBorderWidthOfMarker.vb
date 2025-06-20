Imports System.IO
Imports System.Text
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace SetBorderWidthOfMarker

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SetBorderWidthOfMarker.xlsx")

            ' Get the first chart in the first worksheet of the workbook
            Dim chart As Chart = workbook.Worksheets(0).Charts(0)

            ' Set the marker border width of the first series of the chart to 1.5 points
            chart.Series(0).DataFormat.MarkerBorderWidth = 1.5 'unit is pt

            ' Set the marker border width of the second series of the chart to 2.5 points
            chart.Series(1).DataFormat.MarkerBorderWidth = 2.5 'unit is pt

            ' Save the modified workbook to a new file
            Dim output As String = "SetBorderWidthOfMarker_out.xlsx"
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
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
