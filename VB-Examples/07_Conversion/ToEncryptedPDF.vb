Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls.Pdf.Security
Imports Spire.Xls

Namespace ToEncryptedPdf
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class within a using statement.
            ' The using statement ensures that the workbook object is disposed of properly after use.
            Using workbook As New Workbook()

                ' Load the Excel file from the specified path.
                workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPDF.xlsx")

                ' Set PDF security settings for encryption.
                workbook.ConverterSetting.PdfSecurity.Encrypt("123", "456", PdfPermissionsFlags.Print, PdfEncryptionKeySize.Key128Bit)

                ' Save the workbook to a PDF file format.
                workbook.SaveToFile("sample.pdf", FileFormat.PDF)
            End Using

            'View the document
            ExcelDocViewer("sample.pdf")
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
