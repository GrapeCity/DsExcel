Imports System.Drawing

Namespace Features.Formatting.Fill
    Public Class RectangularGradientFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!A1.Interior.Pattern = Pattern.RectangularGradient
            With TryCast(worksheet.Range!A1.Interior.Gradient, IRectangularGradient)
                .ColorStops(0).Color = Color.Red
                .ColorStops(1).Color = Color.Green
                .Bottom = 0.2
                .Right = 0.3
                .Top = 0.4
                .Left = 0.5
            End With
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
