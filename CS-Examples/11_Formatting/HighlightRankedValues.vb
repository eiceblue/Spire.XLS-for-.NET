Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace HighlightRankedValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new collection of conditional formats for the worksheet
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Define the range of cells to apply conditional formatting to
            xcfs.AddRange(sheet.Range("D2:D10"))

            ' Add a top/bottom conditional format with "Top" type and show top 2 values
            Dim format1 As IConditionalFormat = xcfs.AddTopBottomCondition(TopBottomType.Top, 2)

            ' Set the format type to Top/Bottom
            format1.FormatType = ConditionalFormatType.TopBottom

            ' Set the background color for the conditional format
            format1.BackColor = Color.Red

            ' Create another collection of conditional formats for the worksheet
            Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()

            ' Define the range of cells to apply conditional formatting to
            xcfs1.AddRange(sheet.Range("E2:E10"))

            ' Add a top/bottom conditional format with "Bottom" type and show bottom 2 values
            Dim format2 As IConditionalFormat = xcfs1.AddTopBottomCondition(TopBottomType.Bottom, 2)

            ' Set the format type to Top/Bottom
            format2.FormatType = ConditionalFormatType.TopBottom

            ' Set the background color for the conditional format
            format2.BackColor = Color.ForestGreen

            ' Specify the filename to save the modified workbook
            Dim result As String = "Result-HighlightTopAndBottomRankedValues.xlsx"

            ' Save the workbook to the specified file path using Excel 2013 format
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
