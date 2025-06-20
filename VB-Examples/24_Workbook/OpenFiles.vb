Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO
Imports System.Text

Namespace OpenFiles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Declare file paths for different types of files
            Dim filepath As String = "..\..\..\..\..\..\Data\ExcelSample_N1.xlsx" ' Path to Excel file
            Dim filepath97 As String = "..\..\..\..\..\..\Data\ExcelSample97_N.xls" ' Path to Excel 97-2003 file
            Dim filepathXml As String = "..\..\..\..\..\..\Data\OfficeOpenXML_N.xml" ' Path to XML file
            Dim filepathCsv As String = "..\..\..\..\..\..\Data\CSVSample_N.csv" ' Path to CSV file

            ' Create a StringBuilder object to store messages
            Dim builder As New StringBuilder()

            ' Open and load workbook from the specified Excel file path
            Dim workbook1 As New Workbook()
            workbook1.LoadFromFile(filepath)
            builder.AppendLine("Workbook opened using file path successfully!")

            ' Open and load workbook from the specified file stream
            Dim stream As New FileStream(filepath, FileMode.Open)
            Dim workbook2 As New Workbook()
            workbook2.LoadFromStream(stream)
            builder.AppendLine("Workbook opened using file stream successfully!")
            stream.Dispose()

            ' Open and load an Excel 97-2003 workbook
            Dim wbExcel97 As New Workbook()
            wbExcel97.LoadFromFile(filepath97, ExcelVersion.Version97to2003)
            builder.AppendLine("Microsoft Excel 97 - 2003 workbook opened successfully!")

            ' Open and load an XML file as a workbook
            Dim wbXML As New Workbook()
            wbXML.LoadFromXml(filepathXml)
            builder.AppendLine("XML file opened successfully!")

            ' Open and load a CSV file as a workbook, specifying delimiter and start row/column
            Dim wbCSV As New Workbook()
            wbCSV.LoadFromFile(filepathCsv, ",", 1, 1)
            builder.AppendLine("CSV file opened successfully!")

            ' Write the contents of the StringBuilder to a text file
            Dim output As String = "OpenFiles_out.txt"
            File.WriteAllText(output, builder.ToString())

            ' Release the resources used by the workbook1
            workbook1.Dispose()

            ' Release the resources used by the workbook2
            workbook2.Dispose()

            ' Release the resources used by the wbExcel97
            wbExcel97.Dispose()

            ' Release the resources used by the wbXML
            wbXML.Dispose()

            ' Release the resources used by the wbCSV
            wbCSV.Dispose()

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
