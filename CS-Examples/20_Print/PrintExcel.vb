Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Drawing.Printing

Namespace PrintExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PrintExcel.xlsx")

            ' Get the PrinterSettings object from the workbook's PrintDocument
            Dim settings As PrinterSettings = workbook.PrintDocument.PrinterSettings

            ' Set the starting page number for printing to 0 (first page)
            settings.FromPage = 0

            ' Set the ending page number for printing to 1 (second page)
            settings.ToPage = 1

            ' Print the workbook's contents using the configured printer settings
            workbook.PrintDocument.Print()

            ' Release the resources used by the workbook
            workbook.Dispose()
        End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
