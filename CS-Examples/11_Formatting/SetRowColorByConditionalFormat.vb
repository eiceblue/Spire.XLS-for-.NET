Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections
Imports Spire.Xls.Core

Namespace SetRowColorByConditionalFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the range of cells that contain data in the worksheet
            Dim dataRange As CellRange = sheet.AllocatedRange

            ' Add a new conditional formats collection to the worksheet
            Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            ' Add the data range to the conditional formats collection
            xcfs.AddRange(dataRange)
            ' Add a new conditional format to the collection
            Dim format1 As IConditionalFormat = xcfs.AddCondition()
            ' Set the formula for the conditional format (highlight even rows)
            format1.FirstFormula = "=MOD(ROW(),2)=0"
            ' Set the format type to Formula
            format1.FormatType = ConditionalFormatType.Formula
            ' Set the background color to light sea green
            format1.BackColor = Color.LightSeaGreen

            ' Add another conditional formats collection to the worksheet
            Dim xcfs1 As XlsConditionalFormats = sheet.ConditionalFormats.Add()
            ' Add the data range to the second conditional formats collection
            xcfs1.AddRange(dataRange)
            ' Add a new conditional format to the second collection
            Dim format2 As IConditionalFormat = xcfs.AddCondition()
            ' Set the formula for the conditional format (highlight odd rows)
            format2.FirstFormula = "=MOD(ROW(),2)=1"
            ' Set the format type to Formula
            format2.FormatType = ConditionalFormatType.Formula
            ' Set the background color to yellow
            format2.BackColor = Color.Yellow

            ' Save the workbook to a new file with the name "Result-SetRowColorWithConditionalFormatting.xlsx"
            Dim result As String = "Result-SetRowColorWithConditionalFormatting.xlsx"
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
