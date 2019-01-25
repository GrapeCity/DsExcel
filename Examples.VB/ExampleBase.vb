Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Public MustInherit Class ExampleBase
    Public Sub New()
    End Sub
    Public Overridable ReadOnly Property ID As String
        Get
            Return [GetType]().FullName
        End Get
    End Property
    Public ReadOnly Property Code As String
        Get
            Return GetExampleCode()
        End Get
    End Property
    Public Overridable ReadOnly Property CanDownload As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overridable ReadOnly Property ShowViewer As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overridable ReadOnly Property ShowScreenshot As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overridable ReadOnly Property ShowCode As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overridable ReadOnly Property SavePdf As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overridable ReadOnly Property SaveCsv As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overridable ReadOnly Property HasTemplate As Boolean
        Get
            Return False
        End Get
    End Property
    Friend Property UserAgent As String
    Public Function GetTemplateStream() As Stream
        Return GetResourceStream("" & TemplateName)
    End Function
    Public Function GetResourceStream(resourceName As String) As Stream
        If String.IsNullOrEmpty(resourceName) Then
            Return Nothing
        End If
        ' jack updates resource name after changing assembly name to GrapeCity.Documents.Excel
        resourceName = resourceName.Replace("GrapeCity.Documents.Excel", "GrapeCity.Documents.Spread")
        Dim resource As String = "GrapeCity.Documents.Excel.Examples.VB." & resourceName.Replace("\\", ".")
        Dim assembly = [GetType]().GetTypeInfo().Assembly
        Return assembly.GetManifestResourceStream(resource)
    End Function
    Public Overridable ReadOnly Property TemplateName As String
        Get
            Return Nothing
        End Get
    End Property
    Public Overridable ReadOnly Property IsViewReadOnly As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overridable ReadOnly Property IsUpdate As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overridable ReadOnly Property IsNew As Boolean
        Get
            Return False
        End Get
    End Property
    Protected Overridable ReadOnly Property NameResKey As String
        Get
            Return [GetType]().Name & ".Name"
        End Get
    End Property
    Protected Overridable ReadOnly Property DescripResKey As String
        Get
            Return [GetType]().Name & ".Descrip"
        End Get
    End Property
    Protected ReadOnly Property CurrentDirectory As String
        Get
            Return Directory.GetCurrentDirectory()
        End Get
    End Property
    Public Sub ExecuteExample(workbook As Workbook, userAgents As String())
        BeforeExecute(workbook, userAgents)
        Execute(workbook)
        AfterExecute(workbook, userAgents)
    End Sub
    Protected Overridable Sub BeforeExecute(workbook As Workbook, userAgents As String())
    End Sub
    Public Overridable Sub Execute(workbook As Workbook)
    End Sub
    Protected Overridable Sub AfterExecute(workbook As Workbook, userAgents As String())
        If AgentIsMac(userAgents) Then
            workbook.Calculate() ' ensure that all cached values can be saved in excel file, so number can display the file correctly even if the formulas are not supported in number.
        End If
    End Sub
    Public Overridable ReadOnly Property IsContainedInTree As Boolean
        Get
            Return True
        End Get
    End Property
    Private Function GetExampleCode() As String
        'INSTANT VB NOTE: The variable code was renamed since Visual Basic does not handle local variables named the same as class members well:
        Dim code_Renamed As String = My.Resources.StringResource.ResourceManager.GetString([GetType]().FullName)
        If Not String.IsNullOrWhiteSpace(code_Renamed) Then
            code_Renamed = Regex.Replace(code_Renamed, "[" & vbCrLf & "][^" & vbCrLf & "]\s{8}", vbLf)
        End If
        If SavePdf Then
            code_Renamed &= vbCrLf & "   //save to an pdf file"
            code_Renamed &= String.Format(vbCrLf & "   workbook.Save(""{0}.pdf"", SaveFileFormat.Pdf);", GetShortID())
        ElseIf SaveCsv Then
            code_Renamed &= vbCrLf & "   //save to an csv file"
            code_Renamed &= String.Format(vbCrLf & "   workbook.Save(""{0}.csv"", SaveFileFormat.Csv);", GetShortID())
        ElseIf CanDownload Then
            code_Renamed &= vbCrLf & "   //save to an excel file"
            code_Renamed &= String.Format(vbCrLf & "   workbook.Save(""{0}.xlsx"");", GetShortID())
        End If
        Return code_Renamed
    End Function
    Public Function GetShortID() As String
        Return ID.Substring(ID.LastIndexOf(".") + 1)
    End Function
    Public ReadOnly Property ScreenshotBase64 As String
        Get
            If ShowScreenshot Then
                'INSTANT VB NOTE: The variable id was renamed since Visual Basic does not handle local variables named the same as class members well:
                Dim id_Renamed = [GetType]().FullName
                Dim stream As Stream = GetResourceStream("Screenshots." & id_Renamed & ".png")
                Return ReadStreamToBase64(stream)
            End If
            Return Nothing
        End Get
    End Property
    Public Overridable Function GetNameByCulture(culture As String) As String
        Return My.Resources.StringResource.ResourceManager.GetString(NameResKey, New System.Globalization.CultureInfo(culture))
    End Function
    Public Overridable Function GetDescriptionByCulture(culture As String) As String
        Return My.Resources.StringResource.ResourceManager.GetString(DescripResKey, New System.Globalization.CultureInfo(culture))
    End Function
    Protected Function AgentIsMac(userAgents As String()) As Boolean
        If userAgents.Length > 0 AndAlso userAgents(0).ToLower().Contains("macintosh") Then
            Return True
        End If
        Return False
    End Function
    Private Function ReadStreamToBase64(input As Stream) As String
        Using ms As New MemoryStream()
            input.CopyTo(ms)
            Return "data:image/png;base64," & Convert.ToBase64String(ms.ToArray())
        End Using
    End Function
End Class
Public Class FolderExample
    Inherits ExampleBase
    Private _children As List(Of ExampleBase) = Nothing
    Private _namespace As String
    Public Sub New(ns As String)
        _namespace = ns
    End Sub
    Public Overrides ReadOnly Property ID As String
        Get
            Return _namespace
        End Get
    End Property
    Protected Overrides ReadOnly Property NameResKey As String
        Get
            Dim shortName As String = _namespace.Substring(_namespace.LastIndexOf(".") + 1)
            Return shortName & ".Name"
        End Get
    End Property
    Protected Overrides ReadOnly Property DescripResKey As String
        Get
            Dim shortName As String = _namespace.Substring(_namespace.LastIndexOf(".") + 1)
            Return shortName & ".Descrip"
        End Get
    End Property
    Public ReadOnly Property Children As ExampleBase()
        Get
            If _children Is Nothing Then
                _children = GetChildren()
            End If
            Return _children.ToArray()
        End Get
    End Property
    Private Function GetChildren() As List(Of ExampleBase)
        'INSTANT VB NOTE: The variable children was renamed since Visual Basic does not handle local variables named the same as class members well:
        Dim children_Renamed As New List(Of ExampleBase)()
        Dim types() As Type = AssemblyUtility.GetTypesRecursively(_namespace)
        Dim subNS As New HashSet(Of String)()
        For Each type In types
            If type.Namespace = _namespace Then
                Dim child As ExampleBase = TryCast(Activator.CreateInstance(type), ExampleBase)
                If child.IsContainedInTree Then
                    children_Renamed.Add(child)
                End If
            ElseIf Not subNS.Contains(type.Namespace) Then
                Dim ends As String = type.Namespace.Substring(_namespace.Length + 1)
                If Not String.IsNullOrEmpty(ends) Then
                    Dim nsItems = ends.Split("."c)
                    Dim currentNS = _namespace & "." & nsItems(0)
                    If Not subNS.Contains(currentNS) Then
                        children_Renamed.Add(New FolderExample(currentNS))
                        subNS.Add(currentNS)
                    End If
                    subNS.Add(type.Namespace)
                End If
            End If
        Next type
        children_Renamed.Sort(New ExampleComparer())
        Return children_Renamed
    End Function
    Public Function FindExample(id As String) As ExampleBase
        Return FindExample(Me, id)
    End Function
    Private Function FindExample(example As ExampleBase, id As String) As ExampleBase
        If example.ID = id Then
            Return example
        End If
        Dim folderExample As FolderExample = TryCast(example, FolderExample)
        If folderExample IsNot Nothing Then
            For Each child In folderExample.Children
                Dim result As ExampleBase = FindExample(child, id)
                If result IsNot Nothing Then
                    Return result
                End If
            Next child
        End If
        Return Nothing
    End Function
    Public Overrides ReadOnly Property IsNew As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overrides ReadOnly Property IsUpdate As Boolean
        Get
            Return IsUpdateRecursive(Me)
        End Get
    End Property
    Private Function IsUpdateRecursive(example As ExampleBase) As Boolean
        If TypeOf example Is FolderExample Then
            Dim childFolderExample As FolderExample = TryCast(example, FolderExample)
            For Each item In childFolderExample.Children
                If item.IsUpdate OrElse item.IsNew Then
                    Return True
                End If
                If IsUpdateRecursive(item) Then
                    Return True
                End If
            Next item
        ElseIf example.IsUpdate OrElse example.IsNew Then
            Return True
        End If
        Return False
    End Function
End Class
Public Module AssemblyUtility
    Private _assembly As Assembly = Nothing
    Private _types As List(Of Type) = Nothing
    Private _exampleBaseType As Type = GetType(ExampleBase)
    Private _folderExampleType As Type = GetType(FolderExample)
    Sub New()
        _assembly = GetType(Examples).GetTypeInfo.Assembly
        _types = New List(Of Type)(_assembly.GetTypes)
        _types.Remove(_folderExampleType)
    End Sub
    Public Function GetTypesRecursively(ns As String) As Type()
        Return _types.FindAll(Function(type) type.Namespace IsNot Nothing AndAlso type.Namespace.StartsWith(ns) AndAlso type.GetTypeInfo().BaseType Is _exampleBaseType).ToArray()
    End Function
End Module
