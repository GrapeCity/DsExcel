Imports System.Drawing

Namespace Features.Formatting.Fill
    Public Class PatternFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            With worksheet.Range!A1.Interior
                .Pattern = Pattern.LightDown
                .Color = Color.Pink
                .PatternColorIndex = 5
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
