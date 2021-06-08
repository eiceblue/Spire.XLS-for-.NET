Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace EmbedNoninstalledFonts
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\EmbedNoninstalledFonts.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first chart
			Dim chart As Chart = sheet.Charts(0)

			'Load the font file from disk
			workbook.CustomFontFilePaths = New String() { "..\..\..\..\..\..\Data\PT_Serif-Caption-Web-Regular.ttf" }
			Dim result As System.Collections.Hashtable = workbook.GetCustomFontParsedResult()

			Dim valueList As New ArrayList(result.Values)

			'Apply the font for PrimaryValueAxis of chart
			chart.PrimaryValueAxis.Font.FontName = TryCast(valueList(0), String)

			'Apply the font for PrimaryCategoryAxis of chart
			chart.PrimaryCategoryAxis.Font.FontName = TryCast(valueList(0), String)

			'Apply the font for the first chartSerie of chart
			Dim chartSerie1 As ChartSerie = chart.Series(0)
			chartSerie1.DataPoints.DefaultDataPoint.DataLabels.FontName = TryCast(valueList(0), String)

			Dim output As String ="Output.pdf"
			'Save and Launch
			workbook.SaveToFile(output, Spire.Xls.FileFormat.PDF)
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
