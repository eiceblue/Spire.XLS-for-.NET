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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Save the workbook to a file named "result.xls" in Excel 97-2003 format
            workbook.SaveToFile("result.xls", ExcelVersion.Version97to2003)

            ' Save the workbook to a file named "result.xlsx" in Excel 2010 format
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010)

            ' Save the workbook to a file named "result.xlsb" in Excel 2010 binary format (XLSB)
            workbook.SaveToFile("result.xlsb", ExcelVersion.Xlsb2010)

            ' Save the workbook to a file named "result.ods" in OpenDocument Spreadsheet (ODS) format
            workbook.SaveToFile("result.ods", ExcelVersion.ODS)

            ' Save the workbook to a file named "result.pdf" in PDF format
            workbook.SaveToFile("result.pdf", FileFormat.PDF)

            ' Save the workbook to a file named "result.xml" in XML format
            workbook.SaveToFile("result.xml", FileFormat.XML)

            ' Save the workbook to a file named "result.xps" in XPS format
            workbook.SaveToFile("result.xps", FileFormat.XPS)

            ' Release the resources used by the workbook
            workbook.Dispose()
        End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
