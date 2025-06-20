Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace MoreFormulas
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a workbook
			Dim workbook As New Workbook()

			' Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Write text values
			sheet.Columns(0).NumberFormat = "@"
			sheet.Range("A1").Text = "=CEILING.MATH(-2.78, 5, -1)"
			sheet.Range("A2").Text = "=BITOR(23,10)"
			sheet.Range("A3").Text = "=BITAND(23,10)"
			sheet.Range("A4").Text = "=BITLSHIFT(23,2)"
			sheet.Range("A5").Text = "=BITRSHIFT(23,2)"
			sheet.Range("A6").Text = "=FLOOR.MATH(12.758, 2, -1)"
			sheet.Range("A7").Text = "=ISOWEEKNUM(DATE(2012, 1, 1))"
			sheet.Range("A8").Text = "=CEILING.PRECISE(-4.6, 3)"
			sheet.Range("A9").Text = "=ENCODEURL(""https://www.e-iceblue.com"")"
			sheet.Range("A10").Text = "=ISFORMULA(A1)"
			sheet.Range("A11").Text = "=BITXOR(12, 58)"
			' SPIREXLS-5395
			sheet.Range("A12").Text= "=BAHTTEXT(1234)"
			'SPIREXLS-5393
			sheet.Range("A13").Text = "=TEXTBEFORE(""Red riding hood’s, red hood"", ""hood"")"
			'SPIREXLS - 5394
			sheet.Range("A14").Text = "=TEXTSPLIT(A13,"" "", ""."", TRUE)"
			'SPIREXLS-5397
			sheet.Range("A15").Text = "=TEXTAFTER(""Red riding hood’s, red hood"", ""hood"")"
			',SPIREXLS-5396
			sheet.Range("A16").Text = "= ARRAYTOTEXT(A1：B4，0)"
			'SPIREXLS-5471
			sheet.Range("A17").Text = "=ARABIC(""mcmxii"")"
			'SPIREXLS-5478
			sheet.Range("A18").Text = "=BASE(15,2,10)"
			'SPIREXLS-5479
			sheet.Range("A19").Text = "=COMBINA(3,10)"
			'SPIREXLS-5480
			sheet.Range("A20").Text = "=XOR(3>12,2<9,4>6)"
			' Write formulas
			sheet.Range("B1").Formula = "=CEILING.MATH(-2.78, 5, -1)"
			sheet.Range("B2").Formula = "=BITOR(23,10)"
			sheet.Range("B3").Formula = "=BITAND(23,10)"
			sheet.Range("B4").Formula = "=BITLSHIFT(23,2)"
			sheet.Range("B5").Formula = "=BITRSHIFT(23,2)"
			sheet.Range("B6").Formula = "=FLOOR.MATH(12.758, 2, -1)"
			sheet.Range("B7").Formula = "=ISOWEEKNUM(DATE(2012, 1, 1))"
			sheet.Range("B8").Formula = "=CEILING.PRECISE(-4.6, 3)"
			sheet.Range("B9").Formula = "=ENCODEURL(""https://www.e-iceblue.com"")"
			sheet.Range("B10").Formula = "=ISFORMULA(A1)"
			sheet.Range("B11").Formula = "=BITXOR(12, 58)"
			sheet.Range("B12").Formula = "=BAHTTEXT(1234)"
			sheet.Range("B13").Formula = "=TEXTBEFORE(""Red riding hood’s, red hood"", ""hood"")"
			sheet.Range("B14").Formula = "=TEXTSPLIT(A13,"" "", ""."", TRUE)"
			sheet.Range("B15").Formula = "=TEXTAFTER(""Red riding hood’s, red hood"", ""hood"")"
			sheet.Range("B16").Formula = "=ARRAYTOTEXT(A1：B4，0)"
			sheet.Range("B17").Formula = "=ARABIC(""mcmxii"")"
			sheet.Range("B18").Formula = "=BASE(15,2,10)"
			sheet.Range("B19").Formula = "=COMBINA(3,10)"
			sheet.Range("B20").Formula = "=XOR(3>12,2<9,4>6)"
			' Calculate all value
			workbook.CalculateAllValue()

			' Autofit columns in the allocated range of the sheet
			sheet.AllocatedRange.AutoFitColumns()

			' Save to file 
			Dim result As String = "MoreFormulas.xlsx"
			workbook.SaveToFile(result,ExcelVersion.Version2016)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' View the document
			FileViewer(result)

			Me.Close()
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
