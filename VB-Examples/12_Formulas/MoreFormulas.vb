Imports Spire.Xls

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

			sheet.Range("A12").Text = "=MUNIT(3)"
			sheet.Range("A13").Text = "=FLOOR.PRECISE(3.2)"
			sheet.Range("A14").Text = "=CSC(2)"
			sheet.Range("A15").Text = "=IMCSCH(""4+3i"")"
			sheet.Range("A16").Text = "=IMCOSH(""4+3i"")"
			sheet.Range("A17").Text = "=IMSINH(""4+3i"")"
			sheet.Range("A18").Text = "=IMSECH(""4+3i"")"
			FillData(sheet)
			sheet.Range("A19").Text = "=RANK.AVG(11,H1:H5,1)"
			sheet.Range("A20").Text = "=RANK.EQ(18,H1:H5,1)"
			sheet.Range("A21").Text = "=PERCENTILE.INC(H1:H5,0.3)"
			sheet.Range("A22").Text = "=PERCENTILE.EXC(H1:H5,0.3)"
			sheet.Range("A23").Text = "=BINOM.DIST(6, 10, 0.5, FALSE)"
			sheet.Range("A24").Text = "=BINOM.INV(20, 0.5, 0.9)"


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

			sheet.Range("B12").Formula = "=MUNIT(3)"
			sheet.Range("B13").Formula = "=FLOOR.PRECISE(3.2)"
			sheet.Range("B14").Formula = "=CSC(2)"
			sheet.Range("B15").Formula = "=IMCSCH(""4+3i"")"
			sheet.Range("B16").Formula = "=IMCOSH(""4+3i"")"
			sheet.Range("B17").Formula = "=IMSINH(""4+3i"")"
			sheet.Range("B18").Formula = "=IMSECH(""4+3i"")"
			sheet.Range("B19").Formula = "=RANK.AVG(11,H1:H5,1)"
			sheet.Range("B20").Formula = "=RANK.EQ(18,H1:H5,1)"
			sheet.Range("B21").Formula = "=PERCENTILE.INC(H1:H5,0.3)"
			sheet.Range("B22").Formula = "=PERCENTILE.EXC(H1:H5,0.3)"
			sheet.Range("B23").Formula = "=BINOM.DIST(6, 10, 0.5, FALSE)"
			sheet.Range("B24").Formula = "=BINOM.INV(20, 0.5, 0.9)"


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
		End Sub
		Private Sub FillData(ByVal worksheet As Worksheet)
			worksheet.Range("H1").NumberValue = 20
			worksheet.Range("H2").NumberValue = 11
			worksheet.Range("H3").NumberValue = 18
			worksheet.Range("H4").NumberValue = 22
			worksheet.Range("H5").NumberValue = 6
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
