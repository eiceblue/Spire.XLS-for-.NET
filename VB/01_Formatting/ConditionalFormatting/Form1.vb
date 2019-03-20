Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.ConditionalFormatting
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace ConditionalFormatting
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConditionalFormatting.xlsx")

			 Dim sheet As Worksheet = workbook.Worksheets(0)
            sheet.AllocatedRange.RowHeight = 15
            sheet.AllocatedRange.ColumnWidth = 16
            Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add
            xcfs1.AddRange(sheet.Range("A1:D1"))
            Dim cf1 As IConditionalFormat = xcfs1.AddCondition
            cf1.FormatType = ConditionalFormatType.CellValue
            cf1.FirstFormula = "150"
            cf1.Operator = ComparisonOperatorType.Greater
            cf1.FontColor = Color.Red
            cf1.BackColor = Color.LightBlue
            Dim xcfs2 As XlsConditionalFormats = sheet.ConditionalFormats.Add
            xcfs2.AddRange(sheet.Range("A2:D2"))
            Dim cf2 As IConditionalFormat = xcfs2.AddCondition
            cf2.FormatType = ConditionalFormatType.CellValue
            cf2.FirstFormula = "300"
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
            Dim xcfs3 As XlsConditionalFormats = sheet.ConditionalFormats.Add
            xcfs3.AddRange(sheet.Range("A3:D3"))
            Dim cf3 As IConditionalFormat = xcfs3.AddCondition
            cf3.FormatType = ConditionalFormatType.DataBar
            cf3.DataBar.BarColor = Color.CadetBlue
            Dim xcfs4 As XlsConditionalFormats = sheet.ConditionalFormats.Add
            xcfs4.AddRange(sheet.Range("A4:D4"))
            Dim cf4 As IConditionalFormat = xcfs4.AddCondition
            cf4.FormatType = ConditionalFormatType.IconSet
            cf4.IconSet.IconSetType = IconSetType.ThreeTrafficLights1
            Dim xcfs5 As XlsConditionalFormats = sheet.ConditionalFormats.Add
            xcfs5.AddRange(sheet.Range("A5:D5"))
            Dim cf5 As IConditionalFormat = xcfs5.AddCondition
            cf5.FormatType = ConditionalFormatType.ColorScale
            Dim xcfs6 As XlsConditionalFormats = sheet.ConditionalFormats.Add
            xcfs6.AddRange(sheet.Range("A6:D6"))
            Dim cf6 As IConditionalFormat = xcfs6.AddCondition
            cf6.FormatType = ConditionalFormatType.DuplicateValues
            cf6.BackColor = Color.BurlyWood

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
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
