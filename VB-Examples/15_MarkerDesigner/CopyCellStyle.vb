Imports Spire.Xls

Namespace CopyCellStyle

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\MarkerDesigner1.xlsx")

            ' Create a new DataTable named "data"
            Dim dt As New DataTable("data")

            ' Add two columns, "name" of type String and "age" of type Integer, to the DataTable
            dt.Columns.Add(New DataColumn("name", GetType(String)))
            dt.Columns.Add(New DataColumn("age", GetType(Integer)))

            ' Create three DataRow objects
            Dim drName1 As DataRow = dt.NewRow()
            Dim drName2 As DataRow = dt.NewRow()
            Dim drName3 As DataRow = dt.NewRow()

            ' Assign values to the "name" and "age" columns of each DataRow
            drName1("name") = "John"
            drName1("age") = 15

            drName2("name") = "Jess"
            drName2("age") = 22

            drName3("name") = "Alan"
            drName3("age") = 36

            ' Add the DataRows to the DataTable
            dt.Rows.Add(drName1)
            dt.Rows.Add(drName2)
            dt.Rows.Add(drName3)

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add the DataTable with the name "data" to the MarkerDesigner
            workbook.MarkerDesigner.AddDataTable("data", dt)

            ' Apply the marker designer to replace the markers with data from the DataTable
            workbook.MarkerDesigner.Apply()

            ' Define the output file name as "CopyCellStyle.xlsx"
            Dim output As String = "CopyCellStyle.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            ExcelDocViewer(output)
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
