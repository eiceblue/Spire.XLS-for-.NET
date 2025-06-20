Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace NumberStyles
	Partial Public Class Form1
		Inherits Form

		Public Sub New()

			InitializeComponent()

		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\NumberStyles.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text and make it bold for cell B10
            sheet.Range("B10").Text = "NUMBER FORMATTING"
            sheet.Range("B10").Style.Font.IsBold = True

            ' Set the text and apply number formatting "0" to cell B13 and assign a numeric value to cell C13
            sheet.Range("B13").Text = "0"
            sheet.Range("C13").NumberValue = 1234.5678
            sheet.Range("C13").NumberFormat = "0"

            ' Set the text and apply number formatting "0.00" to cell B14 and assign a numeric value to cell C14
            sheet.Range("B14").Text = "0.00"
            sheet.Range("C14").NumberValue = 1234.5678
            sheet.Range("C14").NumberFormat = "0.00"

            ' Set the text and apply number formatting "#,##0.00" to cell B15 and assign a numeric value to cell C15
            sheet.Range("B15").Text = "#,##0.00"
            sheet.Range("C15").NumberValue = 1234.5678
            sheet.Range("C15").NumberFormat = "#,##0.00"

            ' Set the text and apply number formatting "$#,##0.00" to cell B16 and assign a numeric value to cell C16
            sheet.Range("B16").Text = "$#,##0.00"
            sheet.Range("C16").NumberValue = 1234.5678
            sheet.Range("C16").NumberFormat = "$#,##0.00"

            ' Set the text and apply number formatting "0;[Red]-0" to cell B17 and assign a numeric value to cell C17
            sheet.Range("B17").Text = "0;[Red]-0"
            sheet.Range("C17").NumberValue = -1234.5678
            sheet.Range("C17").NumberFormat = "0;[Red]-0"

            ' Set the text and apply number formatting "0.00;[Red]-0.00" to cell B18 and assign a numeric value to cell C18
            sheet.Range("B18").Text = "0.00;[Red]-0.00"
            sheet.Range("C18").NumberValue = -1234.5678
            sheet.Range("C18").NumberFormat = "0.00;[Red]-0.00"

            ' Set the text and apply number formatting "#,##0;[Red]-#,##0" to cell B19 and assign a numeric value to cell C19
            sheet.Range("B19").Text = "#,##0;[Red]-#,##0"
            sheet.Range("C19").NumberValue = -1234.5678
            sheet.Range("C19").NumberFormat = "#,##0;[Red]-#,##0"

            ' Set the text and apply number formatting "#,##0.00;[Red]-#,##0.00" to cell B20 and assign a numeric value to cell C20
            sheet.Range("B20").Text = "#,##0.00;[Red]-#,##0.000"
            sheet.Range("C20").NumberValue = -1234.5678
            sheet.Range("C20").NumberFormat = "#,##0.00;[Red]-#,##0.00"

            ' Set the text and apply number formatting "0.00E+00" to cell B21 and assign a numeric value to cell C21
            sheet.Range("B21").Text = "0.00E+00"
            sheet.Range("C21").NumberValue = 1234.5678
            sheet.Range("C21").NumberFormat = "0.00E+00"

            ' Set the text and apply number formatting "0.00%" to cell B22 and assign a numeric value to cell C22
            sheet.Range("B22").Text = "0.00%"
            sheet.Range("C22").NumberValue = 1234.5678
            sheet.Range("C22").NumberFormat = "0.00%"

            ' Apply a gray background color to cells B13 to B22
            sheet.Range("B13:B22").Style.KnownColor = ExcelColors.Gray25Percent

            ' Auto-fit column width for columns 2 and 3 (B and C)
            sheet.AutoFitColumn(2)
            sheet.AutoFitColumn(3)

            ' Specify the filename for the resulting Excel file
            Dim result As String = "Result-NumberStyles.xlsx"

            ' Save the workbook to the specified filename in Excel
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
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
