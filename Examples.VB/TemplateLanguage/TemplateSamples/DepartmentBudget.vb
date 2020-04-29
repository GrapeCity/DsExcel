Namespace Templates.TemplateSamples
    Public Class DepartmentBudget
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_DepartmentBudget.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_DepartmentBudget.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            ' Friend Class Departments
            '     Public dpt As List(Of Department)
            ' End Class

            ' Friend Class Department
            '     Public name As String
            '     Public mgr As String
            '     Public bud As Double
            '     Public emp As List(Of Employee)
            ' End Class

            ' Friend Class Employee
            '     Public name As String
            '     Public salary As Double
            ' End Class
#End Region

#Region "Init Data"
            Dim departments As New Departments With {.dpt = New List(Of Department)}

            'Department 1
            Dim department1 As New Department With {
                .name = "Marketing",
                .mgr = "Carl Sommerset",
                .bud = 354586
            }

            department1.emp = New List(Of Employee) From {
                New Employee With {
                    .name = "JoeKline",
                    .salary = 49402
                },
                New Employee With {
                    .name = "Lisa Crane",
                    .salary = 81337
                },
                New Employee With {
                    .name = "John Ryes",
                    .salary = 43503
                },
                New Employee With {
                    .name = "Elli Davidson",
                    .salary = 67334
                },
                New Employee With {
                    .name = "Jack Reaze",
                    .salary = 68314
                },
                New Employee With {
                    .name = "Ben Lam",
                    .salary = 44696
                }
            }

            departments.dpt.Add(department1)

            'Department 2
            Dim department2 As New Department With {
                .name = "Sales",
                .mgr = "Kelly Johnson",
                .bud = 237721
            }

            department2.emp = New List(Of Employee) From {
                New Employee With {
                    .name = "Liam Elmerson",
                    .salary = 61892
                },
                New Employee With {
                    .name = "Angela Sanderson",
                    .salary = 38020
                },
                New Employee With {
                    .name = "Blake Schwarz",
                    .salary = 55701
                },
                New Employee With {
                    .name = "Linda Barataz",
                    .salary = 82108
                }
            }

            departments.dpt.Add(department2)

            'Department 3
            Dim department3 As New Department With {
                .name = "Engineering",
                .mgr = "Gina Davis",
                .bud = 624789
            }

            department3.emp = New List(Of Employee) From {
                New Employee With {
                    .name = "Christopher Dean",
                    .salary = 58329
                },
                New Employee With {
                    .name = "Jack Linner",
                    .salary = 63684
                },
                New Employee With {
                    .name = "Cathy Raines",
                    .salary = 73147
                },
                New Employee With {
                    .name = "Scott Ashton",
                    .salary = 77213
                },
                New Employee With {
                    .name = "Larry Wisell",
                    .salary = 72796
                },
                New Employee With {
                    .name = "Bart Ingram",
                    .salary = 50009
                },
                New Employee With {
                    .name = "Wesley Page",
                    .salary = 82378
                },
                New Employee With {
                    .name = "Alan Keyes",
                    .salary = 67105
                },
                New Employee With {
                    .name = "Wilson Musk",
                    .salary = 80128
                }
            }

            departments.dpt.Add(department3)

            'Department 4
            Dim department4 As New Department With {
                .name = "Customer Service",
                .mgr = "Kenneth Smith",
                .bud = 127596
            }

            department4.emp = New List(Of Employee) From {
                New Employee With {
                    .name = "Sherry Meeks",
                    .salary = 38919
                },
                New Employee With {
                    .name = "Sharon Reeves",
                    .salary = 40963
                },
                New Employee With {
                    .name = "Max Devillo",
                    .salary = 47714
                }
            }

            departments.dpt.Add(department4)
#End Region

            'Add data source
            workbook.AddDataSource("ds", departments)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_DepartmentBudget.xlsx"
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
                Return New String() {"xlsx\Template_DepartmentBudget.xlsx"}
            End Get
        End Property

        Friend Class Departments
            Public dpt As List(Of Department)
        End Class

        Friend Class Department
            Public name As String
            Public mgr As String
            Public bud As Double
            Public emp As List(Of Employee)
        End Class

        Friend Class Employee
            Public name As String
            Public salary As Double
        End Class
    End Class

End Namespace
