Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace DetectExcelVersion
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Define an array of file paths
            Dim files() As String = {"..\..\..\..\..\..\Data\ExcelSample97_N.xls", "..\..\..\..\..\..\Data\ExcelSample_N1.xlsx", "..\..\..\..\..\..\Data\ExcelSample_N.xlsb"}

            ' Create a StringBuilder object to store the output
            Dim builder As New StringBuilder()

            ' Iterate over each file path in the array
            For Each file As String In files

                ' Create a new Workbook object
                Dim workbook As New Workbook()

                ' Load the workbook from the specified file
                workbook.LoadFromFile(file)

                ' Get the version of the loaded workbook
                Dim version As ExcelVersion = workbook.Version

                ' Append the version to the StringBuilder
                builder.AppendLine(version.ToString())

                ' Release the resources used by the workbook
                workbook.Dispose()
            Next file

            ' Specify the output file name
            Dim result As String = "DetectExcelVersion_out.txt"

            ' Write the contents of the StringBuilder to the output file
            File.WriteAllText(result, builder.ToString())

            'Launch the file
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
