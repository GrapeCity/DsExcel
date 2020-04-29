Imports System.Data

Namespace Templates.TemplateSamples
    Public Class SalesTracker
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file from resource
            Dim templateFile = GetResourceStream("xlsx\Template_SalesTracker.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim datasource As New DataTable
            With datasource.Columns
                .Add(New DataColumn("ProductName", GetType(String)))
                .Add(New DataColumn("CostPerItem", GetType(Double)))
                .Add(New DataColumn("PercentMarkup", GetType(Double)))
                .Add(New DataColumn("TotalSold", GetType(Integer)))
                .Add(New DataColumn("ShippingCharge", GetType(Integer)))
                .Add(New DataColumn("ShippingCost", GetType(Double)))
                .Add(New DataColumn("Returns", GetType(Integer)))
            End With

            With datasource.Rows
                .Add("Beverages", 10, 1, 15, 10, 5.75, 2)
                .Add("Condiments", 11.5, 0.75, 18, 10, 5.75, 1)
                .Add("Dairy Products", 13, 0.65, 20, 10, 6.25, 0)
                .Add("Confections", 5, 0.9, 50, 5, 3.5, 0)
                .Add("Sea Food", 4, 0.9, 42, 5, 3.25, 3)
            End With
#End Region

            'Init template global settings
            workbook.Names.Add("TemplateOptions.KeepLineSize", "true")

            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName() As String
            Get
                Return "Template_SalesTracker.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property HasTemplate() As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownloadZip() As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\Template_SalesTracker.xlsx"}
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew() As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
