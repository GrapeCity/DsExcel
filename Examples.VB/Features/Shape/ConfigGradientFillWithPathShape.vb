Namespace Features.Shape
    Public Class ConfigGradientFillWithPathShape
        Inherits ExampleBase
        Public Overrides Sub Execute(ByVal workbook As GrapeCity.Documents.Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\WebsiteFlowChart.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Get "Idea"
            With worksheet.Shapes("Idea").Fill
                .TwoColorGradient(GradientStyle.FromCenter, 1)

                .GradientStops(0).Position = 0.33
                .GradientStops(0).Color.RGB = System.Drawing.Color.FromArgb(0, 112, 192)

                .GradientStops(1).Position = 1
                .GradientStops(1).Color.RGB = System.Drawing.Color.White

                'Set gradient path type as "Path"
                .GradientPathType = PathShapeType.Path
            End With

            'Get "Functionality"
            With worksheet.Shapes("Functionality").Fill
                .TwoColorGradient(GradientStyle.FromCenter, 1)

                .GradientStops(0).Position = 0.33
                .GradientStops(0).Color.RGB = System.Drawing.Color.FromArgb(0, 112, 192)

                .GradientStops(1).Position = 1
                .GradientStops(1).Color.RGB = System.Drawing.Color.White

                'Set gradient path type as "Path"
                .GradientPathType = PathShapeType.Path
            End With
        End Sub
        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\WebsiteFlowChart.xlsx"}
            End Get
        End Property
	End Class
End Namespace
