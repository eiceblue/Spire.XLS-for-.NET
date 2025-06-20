Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SetDataDirection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner2.xlsx")

            ' Create a new DataTable named "data"
            Dim dt As New DataTable("data")

            ' Add a DataColumn named "value" of type String to the DataTable
            dt.Columns.Add(New DataColumn("value", GetType(String)))

            ' Create three DataRow objects
            Dim drName1 As DataRow = dt.NewRow()
            Dim drName2 As DataRow = dt.NewRow()
            Dim drName3 As DataRow = dt.NewRow()

            ' Assign values to the "value" column of each DataRow
            drName1("value") = "Text1"
            drName2("value") = "Text2"
            drName3("value") = "Text3"

            ' Add the DataRows to the DataTable
            dt.Rows.Add(drName1)
            dt.Rows.Add(drName2)
            dt.Rows.Add(drName3)

            ' Add the DataTable with the name "data" to the MarkerDesigner
            workbook.MarkerDesigner.AddDataTable("data", dt)

            ' Apply the marker designer to replace the markers with data from the DataTable
            workbook.MarkerDesigner.Apply()

            ' Define the output file name as "SetDataDirection_result.xlsx"
            Dim output As String = "SetDataDirection_result.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(output)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
