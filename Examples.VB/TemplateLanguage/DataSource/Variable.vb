Namespace Templates.DataSource
    Public Class Variable
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_StudentInfo.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_StudentInfo.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            'Public Class StudentInfo
            '    Public name As String
            '    Public address As String
            '    Public family As List(Of Family)
            'End Class
#End Region

#Region "Init Data"
            Dim studentInfos As New List(Of StudentInfo) From {
                New StudentInfo With {
                    .name = "Jane",
                    .address = "101, Halford Avenue, Fremont, CA"
                },
                New StudentInfo With {
                    .name = "Mark",
                    .address = "2005 Klamath Ave APT, Santa Clara, CA"
                },
                New StudentInfo With {
                    .name = "Carol",
                    .address = "1063 E EI Camino Real, Sunnyvale, CA 94087, USA"
                },
                New StudentInfo With {
                    .name = "Liano",
                    .address = "1977 St Lawrence Dr, Santa Clara, CA 95051, USA"
                },
                New StudentInfo With {
                    .name = "Hellen",
                    .address = "3661 Peacock Ct, Santa Clara, CA 95051, USA"
                }
            }

            Dim className = "Class 3"
#End Region

            'Add data source
            workbook.AddDataSource("className", className)
            workbook.AddDataSource("s", studentInfos)

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
                Return "Template_StudentInfo.xlsx"
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
                Return New String() {"xlsx\Template_StudentInfo.xlsx"}
            End Get
        End Property
    End Class
End Namespace
