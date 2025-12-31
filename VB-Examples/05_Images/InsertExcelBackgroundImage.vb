Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace InsertExcelBackgroundImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Open an image. 
			Dim bm As New Bitmap(Image.FromFile("..\..\..\..\..\..\Data\Background.png"))

			'Set the image to be background image of the worksheet.
			sheet.PageSetup.BackgoundImage = bm

			'////////////////Use the following code for netstandard dlls/////////////////////////
'			
'			SkiaSharp.SKBitmap bm = SkiaSharp.SKBitmap.Decode(@"..\..\..\..\..\..\Data\Background.png");
'            sheet.PageSetup.BackgoundImage = bm;
'			

			' Specify the resulting file name.
			Dim result As String = "Result-InsertExcelBackgroundImage.xlsx"

			' Save the modified workbook to a file using Excel 2013 format.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources.
			workbook.Dispose()

			'Launch the MS Excel file.
			ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub
	End Class
End Namespace
