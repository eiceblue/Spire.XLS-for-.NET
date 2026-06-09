Imports Spire.Xls
Imports System
Imports System.Windows.Forms

Namespace LETAndMAPFunctions
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new workbook
            Dim workbook As Workbook = New Workbook()

            ' Get the first sheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets[0]

            ' Set number values for cells
            sheet.Range["A2"].NumberValue = 1
            sheet.Range["A3"].NumberValue = 2
            sheet.Range["A4"].NumberValue = 3
            sheet.Range["B2"].NumberValue = 11
            sheet.Range["B3"].NumberValue = 12
            sheet.Range["B4"].NumberValue = 13

            'Use the LET function
            sheet.Range["C1"].Text = "out"
            sheet.Range["C2"].Formula = "=LET(x, 5, y, 10, x + y)"
            sheet.Range["C3"].Formula = "=LET(a, 1, b, 2, c, 3, d, 4, a+b+c+d)"
            sheet.Range["C4"].Formula = "=LET(outer, LET(inner, 5, inner*2), outer+10)"

            'Use the MAP function
            sheet.Range["C2"].Formula = "=MAP(A2:A4, LAMBDA(x, x*2))"
            sheet.Range["D2"].Formula = "=MAP(A2:A4,LAMBDA(x,x*10+1))"
            sheet.Range["A8"].Formula = "=MAP(A2:B4,C2:D4,LAMBDA(x,y,SUM(x,y)))"

            ' Recalculate all formulas to ensure values are up to date
            sheet.CalculateAllValue()

            ' Save the modified workbook to the specified file using Excel 2010 format
            Dim result As String = @"LETAndMAPFunctions_out.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Dispose of the workbook object to release resources
            workbook.Dispose()

            'View the document
            FileViewer(result)

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

        Private Sub label1_Click(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace
