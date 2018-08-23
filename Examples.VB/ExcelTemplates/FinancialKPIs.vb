Namespace ExcelTemplates
    Public Class FinancialKPIs
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Financial KPIs.xlsx from resource
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet = workbook.ActiveSheet

            'set values
            Dim ProData = {
                {1483550, 121386},
                {0.4336, 0.32},
                {0.1236, -0.0438},
                {12.36, -0.0438}
            }
            Dim ProData1 = {0.14, 0.0029, 0.0361, 0.0361}
            worksheet.Range("E7:F10").Value = ProData
            worksheet.Range("I7:I10").Value = ProData1

            Dim ActiveData = {
                {"4.32 item", "2.00 item"},
                {"72 days", "45 days"},
                {"53 days", "55 days"}
            }
            Dim ActiveData1 = {"0.45 items", "-5 days", "-6 days"}
            worksheet.Range("E12:F14").Value = ActiveData
            worksheet.Range("I12:I14").Value = ActiveData1

            Dim EffData = {
                {0.3446, 0.25},
                {0.5335, 0.11}
            }
            Dim EffData1 = {0.1245, 0.1946}
            worksheet.Range("E16:F17").Value = EffData
            worksheet.Range("I16:I17").Value = EffData1

            Dim LinData = {
                {"0.91:1", "'2:1"},
                {"0.58:1", "'1:1"}
            }
            Dim LinData1 = {"'0.02:1", "'0.03:1"}
            worksheet.Range("E19:F20").Value = LinData
            worksheet.Range("I19:I20").Value = LinData1

            Dim GeaData = {
                {-9.6, 0.85},
                {0.68, 0.5}
            }
            Dim GeaData1 = {6.65, 0.0282}
            worksheet.Range("E22:F23").Value = GeaData
            worksheet.Range("I22:I23").Value = GeaData1

            Dim CasData = {
                {0.0735, 1.2},
                {0.1442, 0.1442}
            }
            Dim CasData1 = {-0.0046, 0.023}
            worksheet.Range("E25:F26").Value = CasData
            worksheet.Range("I25:I26").Value = CasData1
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Financial KPIs.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property HasTemplate As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
