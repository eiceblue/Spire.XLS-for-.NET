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
            ' Create a new Workbook instance
            Dim workbook As New Workbook()

            ' Load the workbook from a specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first TextBox shape from the worksheet and cast it to XlsTextBoxShape
            Dim shape As XlsTextBoxShape = TryCast(sheet.TextBoxes(0), XlsTextBoxShape)

            ' Create a new ExcelFont instance
            Dim font As ExcelFont = workbook.CreateFont()

            ' Set the properties of the font
            font.FontName = "Century Gothic"
            font.Size = 10
            font.IsBold = True
            font.Color = Color.Blue

            ' Set the font for the entire text in the TextBox shape
            CType(New RichText(shape.RichText), RichText).SetFont(0, shape.Text.Length - 1, font)

            ' Set the fill type of the TextBox shape to solid color
            shape.Fill.FillType = ShapeFillType.SolidColor

            ' Set the foreground color of the TextBox shape to a predefined Excel color
            shape.Fill.ForeKnownColor = ExcelColors.BlueGray

            ' Specify the output file name
            Dim result As String = "Result-SetFontAndBackgroundForTextBox.xlsx"

            ' Save the modified workbook to the specified file path in Excel 2013 format
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
