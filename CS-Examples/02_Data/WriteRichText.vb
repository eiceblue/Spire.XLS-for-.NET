Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace WriteRichText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteRichText.xlsx")
            'Gets the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)
            'Creates a new font object.
            Dim fontBold As ExcelFont = workbook.CreateFont()
            'Sets the font to bold.
            fontBold.IsBold = True
            'Creates a new font object.
            Dim fontUnderline As ExcelFont = workbook.CreateFont()
            'Sets the font to have a single underline.
            fontUnderline.Underline = FontUnderlineType.Single
            'Creates a new font object.
            Dim fontItalic As ExcelFont = workbook.CreateFont()
            'Sets the font to italic.
            fontItalic.IsItalic = True
            'Creates a new font object.
            Dim fontColor As ExcelFont = workbook.CreateFont()
            'Sets the font color to green.
            fontColor.KnownColor = ExcelColors.Green
            'Gets the rich text from cell B11.
            Dim richText As RichText = sheet.Range("B11").RichText
            'Sets the text content of the rich text.
            richText.Text = "Bold and underlined and italic and colored text."
            'Applies the bold font to characters 0 to 3 (Bol).
            richText.SetFont(0, 3, fontBold)
            'Applies the underline font to characters 9 to 18 (underlined).
            richText.SetFont(9, 18, fontUnderline)
            'Applies the italic font to characters 24 to 29 (italic).
            richText.SetFont(24, 29, fontItalic)
            'Applies the green font color to characters 35 to 41 (colored).
            richText.SetFont(35, 41, fontColor)
            'Saves the modified workbook to a file with the specified filename and Excel version.
            workbook.SaveToFile("WriteRichText_result.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("WriteRichText_result.xlsx")
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
