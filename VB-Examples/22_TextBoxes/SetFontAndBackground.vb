Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace SetFontAndBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the textbox which will be edited.
			Dim shape As XlsTextBoxShape = TryCast(sheet.TextBoxes(0), XlsTextBoxShape)

			'Set the font and background color for the textbox.
			'Set font.
			Dim font As ExcelFont = workbook.CreateFont()
			'font.IsStrikethrough = true;
			font.FontName = "Century Gothic"
			font.Size = 10
			font.IsBold = True
			font.Color = Color.Blue
			CType(New RichText(shape.RichText), RichText).SetFont(0, shape.Text.Length - 1, font)

			'Set background color
			shape.Fill.FillType = ShapeFillType.SolidColor
			shape.Fill.ForeKnownColor = ExcelColors.BlueGray

			Dim result As String = "Result-SetFontAndBackgroundForTextBox.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

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
