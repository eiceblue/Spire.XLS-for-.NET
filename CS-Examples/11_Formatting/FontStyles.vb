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
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "FontStyles.xlsx" from a specific location
            workbook.LoadFromFile("..\..\..\..\..\..\Data\FontStyles.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the font name for cell B1 to "Comic Sans MS"
            sheet.Range("B1").Style.Font.FontName = "Comic Sans MS"

            ' Set the font name for cells B2 to D2 to "Corbel"
            sheet.Range("B2:D2").Style.Font.FontName = "Corbel"

            ' Set the font name for cells B3 to D7 to "Aleo"
            sheet.Range("B3:D7").Style.Font.FontName = "Aleo"

            ' Set the font size for cell B1 to 45
            sheet.Range("B1").Style.Font.Size = 45

            ' Set the font size for cells B2 to D3 to 25
            sheet.Range("B2:D3").Style.Font.Size = 25

            ' Set the font size for cells B3 to D7 to 12
            sheet.Range("B3:D7").Style.Font.Size = 12

            ' Set the font style of cells B2 to D2 to bold
            sheet.Range("B2:D2").Style.Font.IsBold = True

            ' Set the font style of cells B3 to B7 to underline
            sheet.Range("B3:B7").Style.Font.Underline = FontUnderlineType.Single

            ' Set the font color of cell B1 to CornflowerBlue
            sheet.Range("B1").Style.Font.Color = Color.CornflowerBlue

            ' Set the font color of cells B2 to D2 to CadetBlue
            sheet.Range("B2:D2").Style.Font.Color = Color.CadetBlue

            ' Set the font color of cells B3 to D7 to Firebrick
            sheet.Range("B3:D7").Style.Font.Color = Color.Firebrick

            ' Set the font style of cells B3 to D7 to italic
            sheet.Range("B3:D7").Style.Font.IsItalic = True

            ' Set the strikethrough font style for cell D3
            sheet.Range("D3").Style.Font.IsStrikethrough = True

            ' Set the strikethrough font style for cell D7
            sheet.Range("D7").Style.Font.IsStrikethrough = True

            ' Save the modified workbook to a new file named "FontStyles_output.xlsx" using Excel 2010 format
            Dim result As String = "FontStyles_output.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
