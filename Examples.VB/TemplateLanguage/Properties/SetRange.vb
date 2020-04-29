Imports System.Collections.Generic
Imports System.Data

Namespace Templates.Properties
    Public Class SetRange
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file from resource
            Dim templateFile = GetResourceStream("xlsx\Template_SetRange.xlsx")
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
                .Add("North America", "Chicago", "Consumer Electronics", "Canon EOS 1500D", 98650)
                .Add("North America", "Chicago", "Consumer Electronics", "Haier 394L 4Star", 367050)
                .Add("South America", "Buenos Aires", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 673800)
                .Add("South America", "Santiago", "Consumer Electronics", "Mi LED 40inch", 102905)
                .Add("South America", "Quito", "Consumer Electronics", "Sennheiser HD 4.40-BT", 234459)
                .Add("North America", "Minnesota", "Mobile", "Iphone XR", 1734621)
                .Add("South America", "Quito", "Mobile", "OnePlus 7Pro", 215000)
                .Add("South America", "Quito", "Mobile", "Redmi 7", 276390)
                .Add("South America", "Buenos Aires", "Mobile", "Samsung S9", 896250)
            End With
#End Region

            'Init template global settings
            workbook.Names.Add("TemplateOptions.KeepLineSize", "true")
            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_SetRange.xlsx"
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
                Return New String() {"xlsx\Template_SetRange.xlsx"}
            End Get
        End Property
    End Class
End Namespace
