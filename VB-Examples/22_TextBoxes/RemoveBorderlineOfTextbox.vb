Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace RemoveBorderlineOfTextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()
            workbook.Version = ExcelVersion.Version2013

            ' Get the first worksheet from the workbook and set its name
            Dim sheet As Worksheet = workbook.Worksheets(0)
            sheet.Name = "Remove Borderline"

            ' Add a chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add()

            ' Add a text box to the chart at position (50, 50) with dimensions (100, 600)
            Dim textbox1 As XlsTextBoxShape = TryCast(chart.TextBoxes.AddTextBox(50, 50, 100, 600), XlsTextBoxShape)
            textbox1.Text = "The solution with borderline"

            ' Add another text box to the chart at position (1000, 50) with dimensions (100, 600)
            Dim textbox2 As XlsTextBoxShape = TryCast(chart.TextBoxes.AddTextBox(1000, 50, 100, 600), XlsTextBoxShape)
            textbox2.Text = "The solution without borderline"
            textbox2.Line.Weight = 0 ' Set the line weight of the text box to 0 to remove the border

            ' Specify the file name for the resulting Excel file
            Dim result As String = "Result-RemoveBorderlineOfTextbox.xlsx"

            ' Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
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
