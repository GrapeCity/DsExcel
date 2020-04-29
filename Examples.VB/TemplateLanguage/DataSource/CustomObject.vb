Imports System.Collections.Generic
Namespace Templates.DataSource
    Public Class CustomObject
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file from resource
            Dim templateFile = GetResourceStream("xlsx\Template_SalesDataGroup.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            ' Public Class SalesData
            '     Public Sales As List(Of SalesRecord)
            ' End Class

            ' Public Class SalesRecord
            '     Public Area As String
            '     Public City As String
            '     Public Category As String
            '     Public Name As String
            '     Public Revenue As Double
            ' End Class
#End Region

#Region "Init Data"
            Dim salesData As New List(Of SalesRecord) From {
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Chicago",
                    .Category = "Consumer Electronics",
                    .Name = "Bose 785593-0050",
                    .Revenue = 92800
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "New York",
                    .Category = "Consumer Electronics",
                    .Name = "Bose 785593-0050",
                    .Revenue = 92800
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Santiago",
                    .Category = "Consumer Electronics",
                    .Name = "Bose 785593-0050",
                    .Revenue = 19550
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Chicago",
                    .Category = "Consumer Electronics",
                    .Name = "Canon EOS 1500D",
                    .Revenue = 98650
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Minnesota",
                    .Category = "Consumer Electronics",
                    .Name = "Canon EOS 1500D",
                    .Revenue = 89110
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Santiago",
                    .Category = "Consumer Electronics",
                    .Name = "Canon EOS 1500D",
                    .Revenue = 459000
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Chicago",
                    .Category = "Consumer Electronics",
                    .Name = "Haier 394L 4Star",
                    .Revenue = 367050
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Quito",
                    .Category = "Consumer Electronics",
                    .Name = "Haier 394L 4Star",
                    .Revenue = 729100
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Santiago",
                    .Category = "Consumer Electronics",
                    .Name = "Haier 394L 4Star",
                    .Revenue = 578900
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Fremont",
                    .Category = "Consumer Electronics",
                    .Name = "IFB 6.5 Kg FullyAuto",
                    .Revenue = 904930
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Buenos Aires",
                    .Category = "Consumer Electronics",
                    .Name = "IFB 6.5 Kg FullyAuto",
                    .Revenue = 673800
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Medillin",
                    .Category = "Consumer Electronics",
                    .Name = "IFB 6.5 Kg FullyAuto",
                    .Revenue = 82910
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Chicago",
                    .Category = "Consumer Electronics",
                    .Name = "Mi LED 40inch",
                    .Revenue = 550010
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Minnesota",
                    .Category = "Consumer Electronics",
                    .Name = "Mi LED 40inch",
                    .Revenue = 1784702
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Santiago",
                    .Category = "Consumer Electronics",
                    .Name = "Mi LED 40inch",
                    .Revenue = 102905
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Chicago",
                    .Category = "Consumer Electronics",
                    .Name = "Sennheiser HD 4.40-BT",
                    .Revenue = 178100
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Quito",
                    .Category = "Consumer Electronics",
                    .Name = "Sennheiser HD 4.40-BT",
                    .Revenue = 234459
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Minnesota",
                    .Category = "Mobile",
                    .Name = "Iphone XR",
                    .Revenue = 1734621
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Santiago",
                    .Category = "Mobile",
                    .Name = "Iphone XR",
                    .Revenue = 109300
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Chicago",
                    .Category = "Mobile",
                    .Name = "OnePlus 7Pro",
                    .Revenue = 499100
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Quito",
                    .Category = "Mobile",
                    .Name = "OnePlus 7Pro",
                    .Revenue = 215000
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Minnesota",
                    .Category = "Mobile",
                    .Name = "Redmi 7",
                    .Revenue = 81650
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Quito",
                    .Category = "Mobile",
                    .Name = "Redmi 7",
                    .Revenue = 276390
                },
                New SalesRecord With {
                    .Area = "North America",
                    .City = "Minnesota",
                    .Category = "Mobile",
                    .Name = "Samsung S9",
                    .Revenue = 896250
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Buenos Aires",
                    .Category = "Mobile",
                    .Name = "Samsung S9",
                    .Revenue = 896250
                },
                New SalesRecord With {
                    .Area = "South America",
                    .City = "Quito",
                    .Category = "Mobile",
                    .Name = "Samsung S9",
                    .Revenue = 716520
                }
            }
#End Region
            Dim datasource = New SalesData With {.Sales = salesData}

            'Init template global settings
            workbook.Names.Add("TemplateOptions.KeepLineSize", "true")
            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_SalesDataGroup.xlsx"
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
                Return New String() {"xlsx\Template_SalesDataGroup.xlsx"}
            End Get
        End Property
        Public Overrides ReadOnly Property Refs As String()
            Get
                Return New String() {"SalesData", "SalesRecord"}
            End Get
        End Property
    End Class

    Public Class SalesData
        Public Sales As List(Of SalesRecord)
    End Class

    Public Class SalesRecord
        Public Area As String
        Public City As String
        Public Category As String
        Public Name As String
        Public Revenue As Double
    End Class
End Namespace
