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
			'Pathes of files
			Dim filepath As String = "..\..\..\..\..\..\Data\ExcelSample_N1.xlsx"
			Dim filepath97 As String = "..\..\..\..\..\..\Data\ExcelSample97_N.xls"
			Dim filepathXml As String = "..\..\..\..\..\..\Data\OfficeOpenXML_N.xml"
			Dim filepathCsv As String = "..\..\..\..\..\..\Data\CSVSample_N.csv"

			'Create string builder
			Dim builder As New StringBuilder()

			'1. Load file by file path
			'Create a workbook
			Dim workbook1 As New Workbook()
			'Load the document from disk
			workbook1.LoadFromFile(filepath)
			builder.AppendLine("Workbook opened using file path successfully!")

			'2. Load file by file stream
			Dim stream As New FileStream(filepath, FileMode.Open)
			'Create a workbook
			Dim workbook2 As New Workbook()
			'Load the document from disk
			workbook2.LoadFromStream(stream)
			builder.AppendLine("Workbook opened using file stream successfully!")
			stream.Dispose()

			'3. Open Microsoft Excel 97 - 2003 file
			Dim wbExcel97 As New Workbook()
			wbExcel97.LoadFromFile(filepath97, ExcelVersion.Version97to2003)
			builder.AppendLine("Microsoft Excel 97 - 2003 workbook opened successfully!")

			'4. Open xml file
			Dim wbXML As New Workbook()
			wbXML.LoadFromXml(filepathXml)
			builder.AppendLine("XML file opened successfully!")

			'5. Open csv file
			Dim wbCSV As New Workbook()
			wbCSV.LoadFromFile(filepathCsv, ",", 1, 1)
			builder.AppendLine("CSV file opened successfully!")

			'Save to txt file
			Dim output As String = "OpenFiles_out.txt"
			File.WriteAllText(output, builder.ToString())

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
