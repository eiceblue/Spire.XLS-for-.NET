Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ApplyMultipleFontsInSingleCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Gets the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Creates a new font object.
            Dim font1 As ExcelFont = workbook.CreateFont()
            'Sets the font color to light blue.
            font1.KnownColor = ExcelColors.LightBlue
            'Sets the font to bold.
            font1.IsBold = True
            'Sets the font size to 10.
            font1.Size = 10

            'Creates a new font object.
            Dim font2 As ExcelFont = workbook.CreateFont()
            'Sets the font color to red.
            font2.KnownColor = ExcelColors.Red
            'Sets the font to bold.
            font2.IsBold = True
            'Sets the font to italic.
            font2.IsItalic = True
            'Sets the font name to Times New Roman.
            font2.FontName = "Times New Roman"
            'Sets the font size to 11.
            font2.Size = 11

            'Gets the rich text from cell H5.
            Dim richText As RichText = sheet.Range("H5").RichText
            'Sets the text content of the rich text.
            richText.Text = "This document was created with Spire.XLS for .NET."
            'Applies the font1 to characters 0 to 29.
            richText.SetFont(0, 29, font1)
            'Applies the font2 to characters 31 to 48.
            richText.SetFont(31, 48, font2)
            'Specifies the filename for the resulting Excel file.
            Dim result As String = "Result-ApplyMultipleFontsInSingleCell.xlsx"

            'Saves the modified workbook to a file with the specified filename and Excel version.
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
