Imports System.Collections.Generic
Imports System.Data

Namespace Templates.Properties
    Public Class SetSortType
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_SetSortType.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim datasource As New DataTable
            With datasource.Columns
                .Add(New DataColumn("Area", GetType(String)))
                .Add(New DataColumn("City", GetType(String)))
                .Add(New DataColumn("Category", GetType(String)))
                .Add(New DataColumn("Name", GetType(String)))
                .Add(New DataColumn("Revenue", GetType(Double)))
            End With

            With datasource.Rows
                .Add("North America", "Chicago", "Consumer Electronics", "Bose 785593-0050", 92800)
                .Add("North America", "New York", "Consumer Electronics", "Bose 785593-0050", 92800)
                .Add("South America", "Santiago", "Consumer Electronics", "Bose 785593-0050", 19550)
                .Add("North America", "Chicago", "Consumer Electronics", "Canon EOS 1500D", 98650)
                .Add("North America", "Minnesota", "Consumer Electronics", "Canon EOS 1500D", 89110)
                .Add("South America", "Santiago", "Consumer Electronics", "Canon EOS 1500D", 459000)
                .Add("North America", "Chicago", "Consumer Electronics", "Haier 394L 4Star", 367050)
                .Add("South America", "Quito", "Consumer Electronics", "Haier 394L 4Star", 729100)
                .Add("South America", "Santiago", "Consumer Electronics", "Haier 394L 4Star", 578900)
                .Add("North America", "Fremont", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 904930)
                .Add("South America", "Buenos Aires", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 673800)
                .Add("South America", "Medillin", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 82910)
                .Add("North America", "Chicago", "Consumer Electronics", "Mi LED 40inch", 550010)
                .Add("North America", "Minnesota", "Consumer Electronics", "Mi LED 40inch", 1784702)
                .Add("South America", "Santiago", "Consumer Electronics", "Mi LED 40inch", 102905)
                .Add("North America", "Chicago", "Consumer Electronics", "Sennheiser HD 4.40-BT", 178100)
                .Add("South America", "Quito", "Consumer Electronics", "Sennheiser HD 4.40-BT", 234459)
                .Add("North America", "Minnesota", "Mobile", "Iphone XR", 1734621)
                .Add("South America", "Santiago", "Mobile", "Iphone XR", 109300)
                .Add("North America", "Chicago", "Mobile", "OnePlus 7Pro", 499100)
                .Add("South America", "Quito", "Mobile", "OnePlus 7Pro", 215000)
                .Add("North America", "Minnesota", "Mobile", "Redmi 7", 81650)
                .Add("South America", "Quito", "Mobile", "Redmi 7", 276390)
                .Add("North America", "Minnesota", "Mobile", "Samsung S9", 896250)
                .Add("South America", "Buenos Aires", "Mobile", "Samsung S9", 896250)
                .Add("South America", "Quito", "Mobile", "Samsung S9", 716520)
            End With
#End Region

            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_SetSortType.xlsx"
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
                Return New String() {"xlsx\Template_SetSortType.xlsx"}
            End Get
        End Property
    End Class
End Namespace
