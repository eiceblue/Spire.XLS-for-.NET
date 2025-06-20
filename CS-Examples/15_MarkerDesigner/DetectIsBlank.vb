Imports Spire.Xls

Namespace DetectIsBlank

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

            ' Create a new DataSet object
            Dim ds As New DataSet()

            ' Read XML data from the specified file path into the DataSet
            ds.ReadXml("..\..\..\..\..\..\Data\Data.xml")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add the DataTable from the DataSet with the name "data" to the MarkerDesigner
            workbook.MarkerDesigner.AddDataTable("data", ds.Tables("data"))

            ' Apply the marker designer to replace the markers with data from the DataTable
            workbook.MarkerDesigner.Apply()

            ' Calculate all the formulas in the workbook
            workbook.CalculateAllValue()

            ' Define the output file name as "DetectIsBlank.xlsx"
            Dim output As String = "DetectIsBlank.xlsx"

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
