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
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set marker designer field in cell A1
			sheet.Range("A1").Value = "&=Student.Name"
			sheet.Range("B1").Value = "&=Student.Age"
			Dim list As New List(Of Student)()
			list.Add(New Student("John", 16))
			list.Add(New Student("Mary", 17))
			list.Add(New Student("Lucy", 17))

			'Fill custom object
			workbook.MarkerDesigner.AddParameter("Student", list)
			workbook.MarkerDesigner.Apply()
			workbook.CalculateAllValue()

			'AutoFit
			sheet.AllocatedRange.AutoFitRows()
			sheet.AllocatedRange.AutoFitColumns()

			'Save the document
			Dim output As String = "AddCustomObject.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

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
