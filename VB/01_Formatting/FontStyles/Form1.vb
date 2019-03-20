Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls

Namespace FontStyles
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FontStyles.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set font style
			sheet.Range("B1").Style.Font.FontName = "Comic Sans MS"
			sheet.Range("B2:D2").Style.Font.FontName = "Corbel"
			sheet.Range("B3:D7").Style.Font.FontName = "Aleo"

			'Set font size
			sheet.Range("B1").Style.Font.Size = 45
			sheet.Range("B2:D3").Style.Font.Size = 25
			sheet.Range("B3:D7").Style.Font.Size = 12

			'Set excel cell data to be bold
			sheet.Range("B2:D2").Style.Font.IsBold = True

			'Set excel cell data to be underline
			sheet.Range("B3:B7").Style.Font.Underline = FontUnderlineType.Single

			'set excel cell data color
			sheet.Range("B1").Style.Font.Color = Color.CornflowerBlue
			sheet.Range("B2:D2").Style.Font.Color = Color.CadetBlue
			sheet.Range("B3:D7").Style.Font.Color = Color.Firebrick

			'set excel cell data to be italic
			sheet.Range("B3:D7").Style.Font.IsItalic = True

			'Save and Launch
			workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010)
			ExcelDocViewer(workbook.FileName)
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
