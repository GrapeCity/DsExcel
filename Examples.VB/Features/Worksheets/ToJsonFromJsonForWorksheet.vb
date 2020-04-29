Imports System.IO

Namespace Features.Worksheets
    Public Class ToJsonFromJsonForWorksheet
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'ToJson&FromJson can be used in combination with spread.sheets product:http://spread.grapecity.com/spreadjs/sheets/

            'GrapeCity Documents for Excel import an excel file.
            'Change the path to real source file path.
            Dim source As String = Path.Combine(CurrentDirectory, "source.xlsx")
            workbook.Open(source)

            'Open the same user file
            Dim new_workbook As New Excel.Workbook
            new_workbook.Open(source)

            For Each worksheet In workbook.Worksheets
                'Do any change of worksheet
                '...

                'GrapeCity Documents for Excel export a worksheet to a json string.
                Dim json As String = worksheet.ToJson()
                'Use the json string to initialize spread.sheets product.
                'Product spread.sheets will show the excel file contents.

                'Use spread.sheets product export a json string of worksheet.
                'GrapeCity Documents for Excel use the json string to update content of the corresponding worksheet.
                new_workbook.Worksheets(worksheet.Name).FromJson(json)
            Next

            'GrapeCity Documents for Excel export workbook to an excel file.
            'Change the path to real export file path.
            Dim export As String = Path.Combine(CurrentDirectory, "export.xlsx")
            new_workbook.Save(export)

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
