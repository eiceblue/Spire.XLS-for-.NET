Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace ConditionalFormatRuntime
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConditionalFormatRuntime.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)
			AddComparisonRule1(sheet)
			AddComparisonRule2(sheet)
			AddComparisonRule3(sheet)
			AddComparisonRule4(sheet)

			Dim result As String = "ConditionalFormatRuntime_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
		End Sub
		Private Sub AddComparisonRule1(ByVal sheet As Worksheet)
			'Create conditional formatting rule
			Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs1.AddRange(sheet.Range("A1:D1"))
			Dim cf1 As IConditionalFormat = xcfs1.AddCondition()
			cf1.FormatType = ConditionalFormatType.CellValue
			cf1.FirstFormula = "150"
			cf1.Operator = ComparisonOperatorType.Greater
			cf1.FontColor = Color.Red
			cf1.BackColor = Color.LightBlue
		End Sub
		Private Sub AddComparisonRule2(ByVal sheet As Worksheet)
			Dim xcfs2 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs2.AddRange(sheet.Range("A2:D2"))
			Dim cf2 As IConditionalFormat = xcfs2.AddCondition()
			cf2.FormatType = ConditionalFormatType.CellValue
			cf2.FirstFormula = "500"
			cf2.Operator = ComparisonOperatorType.Less
			'Set border color
			cf2.LeftBorderColor = Color.Pink
			cf2.RightBorderColor = Color.Pink
			cf2.TopBorderColor = Color.DeepSkyBlue
			cf2.BottomBorderColor = Color.DeepSkyBlue
			cf2.LeftBorderStyle = LineStyleType.Medium
			cf2.RightBorderStyle = LineStyleType.Thick
			cf2.TopBorderStyle = LineStyleType.Double
			cf2.BottomBorderStyle = LineStyleType.Double
		End Sub

		Private Sub AddComparisonRule3(ByVal sheet As Worksheet)
			'Create conditional formatting rule
			Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs1.AddRange(sheet.Range("A3:D3"))
			Dim cf1 As IConditionalFormat = xcfs1.AddCondition()
			cf1.FormatType = ConditionalFormatType.CellValue
			cf1.FirstFormula = "300"
			cf1.SecondFormula = "500"
			cf1.Operator = ComparisonOperatorType.Between
			cf1.BackColor = Color.Yellow
		End Sub

		Private Sub AddComparisonRule4(ByVal sheet As Worksheet)
			'Create conditional formatting rule
			Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
			xcfs1.AddRange(sheet.Range("A4:D4"))
			Dim cf1 As IConditionalFormat = xcfs1.AddCondition()
			cf1.FormatType = ConditionalFormatType.CellValue
			cf1.FirstFormula = "100"
			cf1.SecondFormula = "200"
			cf1.Operator = ComparisonOperatorType.NotBetween
			'Set fill pattern type
			cf1.FillPattern = ExcelPatternType.ReverseDiagonalStripe
			'Set foreground color
			cf1.Color = Color.FromArgb(255, 255, 0)

			'Set background color
			cf1.BackColor = Color.FromArgb(0, 255, 255)
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
