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
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteRichText.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim fontBold As ExcelFont = workbook.CreateFont()
			fontBold.IsBold = True

			Dim fontUnderline As ExcelFont = workbook.CreateFont()
			fontUnderline.Underline = FontUnderlineType.Single

			Dim fontItalic As ExcelFont = workbook.CreateFont()
			fontItalic.IsItalic = True

			Dim fontColor As ExcelFont = workbook.CreateFont()
			fontColor.KnownColor = ExcelColors.Green

			Dim richText As RichText = sheet.Range("B11").RichText
			richText.Text = "Bold and underlined and italic and colored text."
			richText.SetFont(0,3,fontBold)
			richText.SetFont(9,18,fontUnderline)
			richText.SetFont(24, 29, fontItalic)
			richText.SetFont(35,41,fontColor)

			workbook.SaveToFile("WriteRichText_result.xlsx",ExcelVersion.Version2013)
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
