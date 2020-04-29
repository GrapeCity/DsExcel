Namespace Features.Charts.Axes
    Public Class ConfigTickMarkLabelAngle
        Inherits ExampleBase
		Public Overrides Sub Execute(ByVal workbook As GrapeCity.Documents.Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\Manufacturing output chart.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes(0)

            Dim category_axis As IAxis = shape.Chart.Axes.Item(AxisType.Category)

            'config tick label's angle
            category_axis.TickLabels.Orientation = -45
        End Sub

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\Manufacturing output chart.xlsx"}
            End Get
        End Property
	End Class
End Namespace

