Imports System.Drawing

Namespace Features.PDFExporting
    Public Class SaveRangeFillToPDF
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim rangeA1B2 As IRange = worksheet.Range("A1:B2")
            rangeA1B2.Merge()
            rangeA1B2.Interior.Pattern = Pattern.LinearGradient

            Dim gradientA1B2 = TryCast(rangeA1B2.Interior.Gradient, ILinearGradient)
            gradientA1B2.ColorStops(0).Color = Color.Red
            gradientA1B2.ColorStops(1).Color = Color.Yellow
            gradientA1B2.Degree = 90

            Dim rangeE1E2 As IRange = worksheet.Range("D1:E2")
            rangeE1E2.Merge()
            rangeE1E2.Interior.Pattern = Pattern.LightDown
            rangeE1E2.Interior.Color = Color.Pink
            rangeE1E2.Interior.PatternColorIndex = 5

            Dim rangeG1H2 As IRange = worksheet.Range("G1:H2")
            rangeG1H2.Merge()
            rangeG1H2.Interior.Pattern = Pattern.RectangularGradient

            Dim gradientG1H2 = TryCast(rangeG1H2.Interior.Gradient, IRectangularGradient)
            gradientG1H2.ColorStops(0).Color = Color.Red
            gradientG1H2.ColorStops(1).Color = Color.Green
            gradientG1H2.Bottom = 0.2
            gradientG1H2.Right = 0.3
            gradientG1H2.Top = 0.4
            gradientG1H2.Left = 0.5
            worksheet.Range("J1:K2").Merge()
            worksheet.Range("J1:K2").Interior.Color = Color.Green
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
