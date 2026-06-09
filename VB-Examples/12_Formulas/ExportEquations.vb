Imports Spire.Xls
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace ExportEquations
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new workbook
            Dim workbook As Workbook = New Workbook()

            ' Load an existing workbook with a pivot table from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExportEquations.xlsx")

            Dim sheet As Worksheet = workbook.Worksheets[0]

            ' Export the first equation in the sheet to MathML format
            Dim mathML As String = sheet.Equations[0].ExportMathML()
            sheet.Range["B9"].Value = "mathML:"
            sheet.Range["B10"].Value = mathML

            ' Export the first equation to LaTeX format
            Dim LaTex As String = sheet.Equations[0].ExportLaTex()
            sheet.Range["B12"].Value = "LaTeX:"
            sheet.Range["B13"].Value = LaTex

            ' Specify the output file name for the result
            Dim outputFile As String = "ExportEquations_out.xlsx"
            Dim outputFile_TXT As String = "ExportEquations_TXT.txt"

            workbook.SaveToFile(outputFile)
            File.WriteAllText(outputFile_TXT, "LaTeX:\t" + LaTex + "\r\nmathML:\t" + mathML)

            ' Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2010)

            ' Dispose of the workbook object to release resources
            workbook.Dispose()

            'View the document
            FileViewer(outputFile)
        End Sub

        Private Sub FileViewer(ByVal fileName As String)
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
