Imports System.Drawing

Namespace Features.Formatting.Fill
    Public Class LinearGradientFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!A1.Interior.Pattern = Pattern.LinearGradient
            With TryCast(worksheet.Range!A1.Interior.Gradient, ILinearGradient)
                .ColorStops(0).Color = Color.Red
                .ColorStops(1).Color = Color.Yellow
                .Degree = 90
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
