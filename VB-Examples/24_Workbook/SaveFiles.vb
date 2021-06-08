Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SaveFiles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			' Save in Excel 97-2003 format
			workbook.SaveToFile("result.xls",ExcelVersion.Version97to2003)

			' Save in Excel2010 xlsx format
			workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010)

			' Save in XLSB format
			workbook.SaveToFile("result.xlsb", ExcelVersion.Xlsb2010)

			' Save in ODS format
			workbook.SaveToFile("result.ods", ExcelVersion.ODS)

			' Save in PDF format
			workbook.SaveToFile("result.pdf", FileFormat.PDF)

			' Save in XML format
			workbook.SaveToFile("result.xml",FileFormat.XML)

			' Save in XPS format
			workbook.SaveToFile("result.xps", FileFormat.XPS)
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
