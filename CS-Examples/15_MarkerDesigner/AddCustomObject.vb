Imports Spire.Xls

Namespace AddCustomObject

	Partial Public Class Form1
		Inherits Form
		Public Class Student
			Friend Sub New(ByVal name As String, ByVal age As Integer)
				Me.Name = name
				Me.Age = age
			End Sub
			Public Property Name() As String
			Public Property Age() As Integer
		End Class

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value of cell A1 to "&=Student.Name"
            sheet.Range("A1").Value = "&=Student.Name"

            ' Set the value of cell B1 to "&=Student.Age"
            sheet.Range("B1").Value = "&=Student.Age"

            ' Create a list of Student objects
            Dim list As New List(Of Student)()
            list.Add(New Student("John", 16))
            list.Add(New Student("Mary", 17))
            list.Add(New Student("Lucy", 17))

            ' Add the "Student" parameter with the list to the MarkerDesigner
            workbook.MarkerDesigner.AddParameter("Student", list)

            ' Apply the marker designer to replace the markers with actual values
            workbook.MarkerDesigner.Apply()

            ' Calculate all the formulas in the workbook
            workbook.CalculateAllValue()

            ' Autofit the rows and columns of the allocated range in the worksheet
            sheet.AllocatedRange.AutoFitRows()
            sheet.AllocatedRange.AutoFitColumns()

            ' Define the output file name as "AddCustomObject.xlsx"
            Dim output As String = "AddCustomObject.xlsx"

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
