Namespace Features.Workbook
    Public Class ToJsonFromJson
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'ToJson&FromJson can be used in combination with spread.sheets product:http://spread.grapecity.com/spreadjs/sheets/
            'GrapeCity Documents for Excel import an excel file.
            'change the path to real source file path.
            Dim source As String = IO.Path.Combine(CurrentDirectory, "source.xlsx")
            workbook.Open(source)

            'GrapeCity Documents for Excel export to a json string.
            Dim jsonstr = workbook.ToJson()

            'use the json string to initialize spread.sheets product.
            'spread.sheets will show the excel file contents.
            'spread.sheets product export a json string.
            'GrapeCity Documents for Excel use the json string to initialize.
            workbook.FromJson(jsonstr)

            'GrapeCity Documents for Excel export workbook to an excel file.
            'change the path to real export file path.
            Dim export As String = IO.Path.Combine(CurrentDirectory, "export.xlsx")

            workbook.Save(export)
        End Sub
        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
