Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace MarkerDesigner
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim sourceData As New Workbook()

            ' Load an existing workbook from the specified file path
            sourceData.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner-DataSample.xls")

            ' Get the first worksheet in the workbook
            Dim sourceSheet As Worksheet = sourceData.Worksheets(0)

            ' Export sheet data to a data table
            Dim dt As DataTable = sourceSheet.ExportDataTable()

            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner.xls")


            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a parameter named "Variable1" with the value 1234.5678 to the MarkerDesigner
            workbook.MarkerDesigner.AddParameter("Variable1", 1234.5678)

            ' Add the DataTable with the name "Country" to the MarkerDesigner
            workbook.MarkerDesigner.AddDataTable("Country", dt)

            ' Apply the marker designer to replace the markers with actual values
            workbook.MarkerDesigner.Apply()

            ' Autofit the rows and columns of the allocated range in the worksheet
            sheet.AllocatedRange.AutoFitRows()
            sheet.AllocatedRange.AutoFitColumns()

            ' Define the output file name as "Output_MarkerDesigner.xlsx"
            Dim result As String = "Output_MarkerDesigner.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2010 format
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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner-DataSample.xls")
			'Initailize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Me.dataGrid1.DataSource = sheet.ExportDataTable()
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub


	End Class
End Namespace
