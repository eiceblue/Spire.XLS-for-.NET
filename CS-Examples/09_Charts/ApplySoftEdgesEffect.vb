Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ApplySoftEdgesEffect

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "ChartSample3.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample3.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Apply a soft edge effect to the chart's chart area with a value of 25
            chart.ChartArea.Shadow.SoftEdge = 25

            ' Specify the output file name as "ApplySoftEdgesEffect.xlsx"
            Dim output As String = "ApplySoftEdgesEffect.xlsx"

            ' Save the modified workbook to the specified file path, using the Excel 2013 format
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
