Imports Spire.Xls
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace ImageHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			' Create a new workbook
			Dim workbook As New Workbook()

			' Load a Workbook from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ImageHeaderFooter.xlsx")

			' Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Load an image from disk
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Logo.png")

			'////////////////Use the following code for netstandard dlls/////////////////////////
'			
'			SkiaSharp.SKBitmap image = SkiaSharp.SKBitmap.Decode((@"..\..\..\..\..\..\Data\Logo.png");
'			

			' Set the image header
			sheet.PageSetup.LeftHeaderImage = image
			sheet.PageSetup.LeftHeader = "&G"

			' Set the image footer
			sheet.PageSetup.CenterFooterImage = image
			sheet.PageSetup.CenterFooter = "&G"

			' Set the view mode of the sheet
			sheet.ViewMode = ViewMode.Layout

			' Specify the file name for the resulting Excel file
			Dim result As String ="Output_ImageHeaderFooter.xlsx"

			' Save the workbook to the specified file in Excel 2013 format
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
