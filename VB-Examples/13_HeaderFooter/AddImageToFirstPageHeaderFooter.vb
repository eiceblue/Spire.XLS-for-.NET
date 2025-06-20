Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Text
Imports System.IO

Namespace AddImageToFirstPageHeaderFooter
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		 Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new workbook
			Dim workbook As New Workbook()

			' Load a Workbook from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AddImageToFirstPageHeaderFooter.xlsx")

			' Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.PageSetup.DifferentFirst = CByte(1)

			' Load an image from disk
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Logo.png")

			' Set the image header
			sheet.PageSetup.FirstLeftHeaderImage = image
			sheet.PageSetup.FirstCenterHeaderImage = image
			sheet.PageSetup.FirstRightHeaderImage = image

			' Set the image footer
			sheet.PageSetup.FirstLeftFooterImage = image
			sheet.PageSetup.FirstCenterFooterImage = image
			sheet.PageSetup.FirstRightFooterImage = image

			' Set the view mode of the sheet
			sheet.ViewMode = ViewMode.Layout

			' Specify the file name for the resulting Excel file
			Dim result As String = "Output_AddImageHeaderFooterToFirstPage.xlsx"

			' Save the workbook to the specified file in Excel 2016 format
			workbook.SaveToFile(result, ExcelVersion.Version2016)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
