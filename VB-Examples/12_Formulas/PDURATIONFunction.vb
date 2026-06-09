Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Spire.Xls

Namespace PDURATIONFunction
	Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)

            'Create workbook instance
            Dim workbook As Workbook = New Workbook()

            'Get first worksheet
            Dim sheet As Worksheet = workbook.Worksheets[0]

            'Set headers
            sheet.Range["A1"].Text = "Financial Calculation"
            sheet.Range["A2"].Text = "Annual Interest Rate"
            sheet.Range["A3"].Text = "Present Value (PV)"
            sheet.Range["A4"].Text = "Future Value (FV)"
            sheet.Range["A5"].Text = "Required Periods (Years)"

            'Set input values
            sheet.Range["B2"].Text = "2.5%"
            sheet.Range["B3"].Text = "2000"
            sheet.Range["B4"].Text = "2200"

            'Calculate years to grow from 2000 to 2200 at 2.5% annual rate using PDURATION function
            sheet.Range["B5"].Formula = "=PDURATION(2.5%,2000,2200)"
            'Format cells
            sheet.Range["B2"].Style.NumberFormat = "0.00%"
            sheet.Range["B3:B4"].Style.NumberFormat = "#,##0"
            sheet.Range["B5"].Style.NumberFormat = "0.00"

            'Auto fit columns
            sheet.AllocatedRange.AutoFitColumns()

            ' Specify the file name for the resulting Excel file
            Dim result As String = "PDURATION_Calculation.xlsx"

            ' Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

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

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
	End Class
End Namespace
