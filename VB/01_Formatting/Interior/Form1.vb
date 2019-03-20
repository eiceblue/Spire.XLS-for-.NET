Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace Interior
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Initialize the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Specify the version
			workbook.Version = ExcelVersion.Version2007

			'Define the number of the colors
			Dim maxColor As Integer = System.Enum.GetValues(GetType(ExcelColors)).Length

			'Create a random object
            Dim random As New Random(10000)

			For i As Integer = 2 To 39
				'Random backKnownColor
				Dim backKnownColor As ExcelColors = CType(random.Next(1, maxColor \ 2), ExcelColors)

				'Add text
				sheet.Range("A1").Text = "Color Name"
				sheet.Range("B1").Text = "Red"
				sheet.Range("C1").Text = "Green"
				sheet.Range("D1").Text = "Blue"

				'Merge the sheet"E1-K1"
				sheet.Range("E1:K1").Merge()
				sheet.Range("E1:K1").Text = "Gradient"
				sheet.Range("A1:K1").Style.Font.IsBold = True
				sheet.Range("A1:K1").Style.Font.Size = 11

				'Set the text of color in sheetA-sheetD
				Dim colorName As String = backKnownColor.ToString()
				sheet.Range(String.Format("A{0}", i)).Text = colorName
				sheet.Range(String.Format("B{0}", i)).NumberValue = workbook.GetPaletteColor(backKnownColor).R
				sheet.Range(String.Format("C{0}", i)).NumberValue = workbook.GetPaletteColor(backKnownColor).G
				sheet.Range(String.Format("D{0}", i)).NumberValue = workbook.GetPaletteColor(backKnownColor).B

				'Merge the sheets
				sheet.Range(String.Format("E{0}:K{0}", i)).Merge()

				'Set the text of sheetE-sheetK
				sheet.Range(String.Format("E{0}:K{0}", i)).Text = colorName

				'Set the interior of the color
				sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.FillPattern = ExcelPatternType.Gradient
				sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.BackKnownColor = backKnownColor
				sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.ForeKnownColor = ExcelColors.White
				sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical
				sheet.Range(String.Format("E{0}:K{0}", i)).Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1
			Next i

			'AutoFit Column
			sheet.AutoFitColumn(1)

			'Save and Launch
			workbook.SaveToFile("output.xlsx",ExcelVersion.Version2013)
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
