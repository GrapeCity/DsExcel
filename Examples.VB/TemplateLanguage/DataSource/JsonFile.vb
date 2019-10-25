Imports System.IO
Imports Newtonsoft.Json

Namespace Templates.DataSource
    Public Class JsonFile
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_FamilyInfo.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_FamilyInfo.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            'Public Class StudentInfos
            '    Public student As List(Of StudentInfo)
            'End Class

            'Public Class StudentInfo
            '    Public name As String
            '    Public address As String
            '    Public family As List(Of Family)
            'End Class

            'Public Class Family
            '    Public father As Guardian
            '    Public mother As Guardian
            'End Class

            'Public Class Guardian
            '    Public name As String
            '    Public occupation As String
            'End Class
#End Region

            'Get data from json file
            Dim jsonText As String
            Using stream = GetResourceStream("Template_FamilyInfo.json"),
                  reader As New StreamReader(stream)
                jsonText = reader.ReadToEnd()
            End Using

            Dim datasource = JsonConvert.DeserializeObject(Of StudentInfos)(jsonText)

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
                Return "Template_FamilyInfo.xlsx"
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
                Return New String() {"xlsx\Template_FamilyInfo.xlsx", "Template_FamilyInfo.json"}
            End Get
        End Property
    End Class

    Public Class StudentInfos
        Public student As List(Of StudentInfo)
    End Class

    Public Class StudentInfo
        Public name As String
        Public address As String
        Public family As List(Of Family)
    End Class

    Public Class Family
        Public father As Guardian
        Public mother As Guardian
    End Class

    Public Class Guardian
        Public name As String
        Public occupation As String
    End Class
End Namespace
