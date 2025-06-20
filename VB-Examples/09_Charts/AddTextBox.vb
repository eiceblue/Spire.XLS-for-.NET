Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports Spire.Xls.Core

Namespace AddTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "AddTextBox.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AddTextBox.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart (index 0) on the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Add a textbox to the chart and assign it to the variable "textbox"
            Dim textbox As ITextBoxLinkShape = chart.Shapes.AddTextBox()

            ' Set the width of the textbox to 1200 units
            textbox.Width = 1200

            ' Set the height of the textbox to 320 units
            textbox.Height = 320

            ' Set the left position of the textbox to 1000 units
            textbox.Left = 1000

            ' Set the top position of the textbox to 480 units
            textbox.Top = 480

            ' Set the text content of the textbox
            textbox.Text = "This is a textbox"

            ' Save the modified workbook to a new Excel file named "Output.xlsx" with the Excel 2010 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Output.xlsx")
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
