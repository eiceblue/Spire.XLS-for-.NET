Imports Spire.Xls
Imports System.IO

Namespace ChartSheetToSVG

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSheet.xlsx")

			'Get the second chartsheet by name
			Dim cs As ChartSheet = workbook.GetChartSheetByName("Chart1")

			'Save to SVG stream
			Dim output As String = "ToSVG.svg"
			Dim fs As New FileStream(String.Format(output), FileMode.Create)
			cs.ToSVGStream(fs)
			fs.Flush()
			fs.Close()

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
