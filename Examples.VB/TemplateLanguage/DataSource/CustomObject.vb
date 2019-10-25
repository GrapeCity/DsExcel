Namespace Templates.DataSource
    Public Class CustomObject
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_SalesDataGroup.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_SalesDataGroup.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            ' Public Class SalesData
            '     Public Records As List(Of SalesRecord)
            ' End Class

            ' Public Class SalesRecord
            '     Public Area As String
            '     Public Salesman As String
            '     Public Product As String
            '     Public ProductType As String
            '     Public Sales As Integer
            ' End Class
#End Region

#Region "Init Data"
            Dim datasource = New SalesData With {
                .Records = New List(Of SalesRecord) From {
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Hellen",
                        .Product = "Apple",
                        .ProductType = "Fruit",
                        .Sales = 120
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Hellen",
                        .Product = "Banana",
                        .ProductType = "Fruit",
                        .Sales = 143
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Hellen",
                        .Product = "Kiwi",
                        .ProductType = "Fruit",
                        .Sales = 322
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Hellen",
                        .Product = "Carrots",
                        .ProductType = "Vegetable",
                        .Sales = 154
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Fancy",
                        .Product = "Carrots",
                        .ProductType = "Vegetable",
                        .Sales = 131
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Fancy",
                        .Product = "Cabbage",
                        .ProductType = "Vegetable",
                        .Sales = 98
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Fancy",
                        .Product = "Potato",
                        .ProductType = "Vegetable",
                        .Sales = 212
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Fancy",
                        .Product = "Apple",
                        .ProductType = "Fruit",
                        .Sales = 102
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Ivan",
                        .Product = "Apple",
                        .ProductType = "Fruit",
                        .Sales = 164
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Ivan",
                        .Product = "Kiwi",
                        .ProductType = "Fruit",
                        .Sales = 213
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Ivan",
                        .Product = "Potato",
                        .ProductType = "Vegetable",
                        .Sales = 56
                    },
                    New SalesRecord With {
                        .Area = "NorthChina",
                        .Salesman = "Ivan",
                        .Product = "Cabbage",
                        .ProductType = "Vegetable",
                        .Sales = 265
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Adam",
                        .Product = "Cabbage",
                        .ProductType = "Vegetable",
                        .Sales = 112
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Adam",
                        .Product = "Carrots",
                        .ProductType = "Vegetable",
                        .Sales = 354
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Adam",
                        .Product = "Banana",
                        .ProductType = "Fruit",
                        .Sales = 277
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Adam",
                        .Product = "Apple",
                        .ProductType = "Fruit",
                        .Sales = 105
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Bob",
                        .Product = "Banana",
                        .ProductType = "Fruit",
                        .Sales = 133
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Bob",
                        .Product = "Cabbage",
                        .ProductType = "Vegetable",
                        .Sales = 252
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Bob",
                        .Product = "Potato",
                        .ProductType = "Vegetable",
                        .Sales = 265
                    },
                    New SalesRecord With {
                        .Area = "SouthChina",
                        .Salesman = "Bob",
                        .Product = "Kiwi",
                        .ProductType = "Fruit",
                        .Sales = 402
                    }
                }
            }

#End Region

            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

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
    End Class

    Public Class SalesData
        Public Records As List(Of SalesRecord)
    End Class

    Public Class SalesRecord
        Public Area As String
        Public Salesman As String
        Public Product As String
        Public ProductType As String
        Public Sales As Integer
    End Class
End Namespace
