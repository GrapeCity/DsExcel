Imports System.Data

Namespace Templates.TemplateSamples
    Public Class FinancialDashboard
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_FinancialDashboard.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim datasource As New DataTable

            With datasource.Columns
                .Add(New DataColumn("season", GetType(String)))
                .Add(New DataColumn("country", GetType(String)))
                .Add(New DataColumn("expect", GetType(Double)))
                .Add(New DataColumn("actual", GetType(Double)))
            End With

            With datasource.Rows
                .Add("2016 Q1", "USA", 236047, 328554)
                .Add("2016 Q2", "USA", 373060, 238136)
                .Add("2016 Q3", "USA", 224132, 300822)
                .Add("2016 Q4", "USA", 269305, 315337)
                .Add("2017 Q1", "USA", 265397, 279008)
                .Add("2017 Q2", "USA", 214079, 206019)
                .Add("2017 Q3", "USA", 370191, 238294)
                .Add("2017 Q4", "USA", 266843, 242323)
                .Add("2016 Q1", "Japan", 350156, 370834)
                .Add("2016 Q2", "Japan", 369399, 247324)
                .Add("2016 Q3", "Japan", 278834, 237385)
                .Add("2016 Q4", "Japan", 264277, 245048)
                .Add("2017 Q1", "Japan", 203006, 295389)
                .Add("2017 Q2", "Japan", 276987, 215804)
                .Add("2017 Q3", "Japan", 330315, 330443)
                .Add("2017 Q4", "Japan", 307477, 262512)
                .Add("2016 Q1", "Korea", 229432, 330368)
                .Add("2016 Q2", "Korea", 321904, 279114)
                .Add("2016 Q3", "Korea", 230496, 219257)
                .Add("2016 Q4", "Korea", 254328, 361880)
                .Add("2017 Q1", "Korea", 272263, 355419)
                .Add("2017 Q2", "Korea", 214079, 231510)
                .Add("2017 Q3", "Korea", 238392, 237430)
                .Add("2017 Q4", "Korea", 294097, 257680)
                .Add("2016 Q1", "China", 238175, 266070)
                .Add("2016 Q2", "China", 202721, 353563)
                .Add("2016 Q3", "China", 253279, 312586)
                .Add("2016 Q4", "China", 211847, 306970)
                .Add("2017 Q1", "China", 369314, 315718)
                .Add("2017 Q2", "China", 201224, 368630)
                .Add("2017 Q3", "China", 239792, 255108)
                .Add("2017 Q4", "China", 271096, 297354)
                .Add("2016 Q1", "India", 236047, 328554)
                .Add("2016 Q2", "India", 373060, 238136)
                .Add("2016 Q3", "India", 224132, 300822)
                .Add("2016 Q4", "India", 269305, 315337)
                .Add("2017 Q1", "India", 265397, 279008)
                .Add("2017 Q2", "India", 214079, 206019)
                .Add("2017 Q3", "India", 370191, 238294)
                .Add("2017 Q4", "India", 266843, 242323)
            End With
#End Region

            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_FinancialDashboard.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property HasTemplate As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownloadZip As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Template_FinancialDashboard.xlsx"}
            End Get
        End Property
    End Class
End Namespace
