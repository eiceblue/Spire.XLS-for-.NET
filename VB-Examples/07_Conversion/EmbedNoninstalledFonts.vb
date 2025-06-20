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
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\EmbedNoninstalledFonts.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Specify the file paths for custom fonts to be embedded in the workbook
            workbook.CustomFontFilePaths = New String() {"..\..\..\..\..\..\Data\PT_Serif-Caption-Web-Regular.ttf"}

            ' Retrieve the parsed result of custom fonts
            Dim result As System.Collections.Hashtable = workbook.GetCustomFontParsedResult()

            ' Extract the font names from the parsed result
            Dim valueList As New ArrayList(result.Values)

            ' Set the font name for the primary value axis of the chart
            chart.PrimaryValueAxis.Font.FontName = TryCast(valueList(0), String)

            ' Set the font name for the primary category axis of the chart
            chart.PrimaryCategoryAxis.Font.FontName = TryCast(valueList(0), String)

            ' Get the first series of the chart
            Dim chartSerie1 As ChartSerie = chart.Series(0)

            ' Set the font name for the data labels of the default data point in the series
            chartSerie1.DataPoints.DefaultDataPoint.DataLabels.FontName = TryCast(valueList(0), String)

            ' Specify the output file name for saving as PDF
            Dim output As String = "Output.pdf"

            ' Save the workbook to a PDF file
            workbook.SaveToFile(output, Spire.Xls.FileFormat.PDF)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
