Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.IO

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


    Public ReadOnly Property CodeVB As String
        Get
            Return GetExampleCodeVB()
        End Get
    End Property

    Public Overridable ReadOnly Property CanDownload As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overridable ReadOnly Property CanDownloadZip As Boolean
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

    Public Overridable ReadOnly Property SavePageInfos As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overridable ReadOnly Property SaveAsImages As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overridable ReadOnly Property SaveWorkbooks As Boolean
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


    Public Overridable ReadOnly Property UsedResources As String()
        Get
            Return Nothing
        End Get
    End Property

    Friend Property UserAgent As String

    Public Function GetTemplateStream() As Stream
        Return GetResourceStream("" & TemplateName)
    End Function

    Protected Function GetResourceStream(resourcePath As String) As Stream
        Dim resource As String = "%assembly%." + resourcePath.Substring(resourcePath.LastIndexOf("\"c) + 1)
        Dim assembly = GetType(Program).GetTypeInfo().Assembly
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

    Public Sub ExecuteExample(workbook As Workbook, userAgents() As String)
        BeforeExecute(workbook, userAgents)
        Execute(workbook)
        AfterExecute(workbook, userAgents)
    End Sub

    Public Sub ExecuteExample(stream As MemoryStream, workbook As Workbook, userAgents() As String)
        BeforeExecute(workbook, userAgents)
        Execute(workbook, stream)
        AfterExecute(workbook, userAgents)
    End Sub

    Protected Overridable Sub BeforeExecute(workbook As Workbook, userAgents() As String)

    End Sub

    Public Overridable Sub Execute(workbook As Workbook)

    End Sub

    Public Overridable Sub Execute(workbook As Workbook, outputStream As MemoryStream)

    End Sub

    Protected Overridable Sub AfterExecute(workbook As Workbook, userAgents() As String)
        If AgentIsMac(userAgents) Then workbook.Calculate() ' ensure that all cached values can be saved in excel file, so number can display the file correctly even if the formulas are not supported in number.
    End Sub

    Public Overridable ReadOnly Property IsContainedInTree As Boolean
        Get
            Return True
        End Get
    End Property

    Public Function GetResourceText(resourceName As String) As String
        Using reader As New StreamReader(GetResourceStream(resourceName))
            Return reader.ReadToEnd()
        End Using
    End Function

    Private Function GetExampleCode() As String
        Dim streamCode As String = Nothing
        If SavePageInfos Then
            streamCode = "   //Note: To use the PrintManager, you should have valid license for GrapeCity Documents for PDF." & vbCrLf & vbCrLf

            streamCode &= "   //create a pdf file stream"
            streamCode &= String.Format(vbCrLf & "   FileStream outputStream = new FileStream(""{0}.pdf"", FileMode.Create);" & vbCrLf & vbCrLf, GetShortID())
        ElseIf SaveAsImages Then
            streamCode = "   //create a png file stream"
            streamCode &= String.Format(vbCrLf & "   FileStream outputStream = new FileStream(""{0}.png"", FileMode.Create);" & vbCrLf & vbCrLf, GetShortID())
        End If

        Dim code As String = My.Resources.CodeResource.ResourceManager.GetString([GetType]().FullName)
        If Not String.IsNullOrWhiteSpace(code) Then code = Regex.Replace(code, "[" & vbCrLf & "][^" & vbCrLf & "]\s{8}", vbLf)

        If SavePdf Then
            code &= vbCrLf & "   //save to a pdf file"
            code &= String.Format(vbCrLf & "   workbook.Save(""{0}.pdf"");", GetShortID())
        ElseIf SavePageInfos OrElse SaveAsImages Then
            code &= vbCrLf & "   //close the pdf stream"
            code &= String.Format(vbCrLf & "   outputStream.Close();")
        ElseIf SaveCsv Then
            code &= vbCrLf & "   //save to a csv file"
            code &= String.Format(vbCrLf & "   workbook.Save(""{0}.csv"");", GetShortID())
        ElseIf CanDownload Then
            code &= vbCrLf & "   //save to an excel file"
            code &= String.Format(vbCrLf & "   workbook.Save(""{0}.xlsx"");", GetShortID())
        End If
        Return streamCode & code
    End Function

    Private Function GetExampleCodeVB() As String
        Dim streamCode As String = Nothing
        If SavePageInfos Then
            streamCode = "   ' Create a pdf file stream"
            streamCode &= String.Format(vbCrLf & "   Dim outputStream = File.Create(""{0}.pdf"")" & vbCrLf & vbCrLf, GetShortID())
        ElseIf SaveAsImages Then
            streamCode = "   ' Create a png file stream"
            streamCode &= String.Format(vbCrLf & "   Dim outputStream = File.Create(""{0}.png"")" & vbCrLf & vbCrLf, GetShortID())
        End If

        Dim code As String = My.Resources.CodeResource_VB.ResourceManager.GetString([GetType]().FullName.Replace("GrapeCity.Documents.Excel.Examples", "GrapeCity.Documents.Excel.Examples.VB"))
        If Not String.IsNullOrWhiteSpace(code) Then code = Regex.Replace(code, "[" & vbCrLf & "][^" & vbCrLf & "]\s{8}", vbLf)

        code = "   ' Create a new Workbook" & Environment.NewLine & "   Dim workbook As New Workbook" & code

        If SavePdf Then
            code &= vbCrLf & "  ' save to a pdf file"
            code &= String.Format(vbCrLf & "   workbook.Save(""{0}.pdf"")", GetShortID())
        ElseIf SaveCsv Then
            code &= vbCrLf & "  ' save to a csv file"
            code &= String.Format(vbCrLf & "   workbook.Save(""{0}.csv"")", GetShortID())
        ElseIf SavePageInfos OrElse SaveAsImages Then
            code &= vbCrLf & "   ' close the pdf stream"
            code &= String.Format(vbCrLf & "   outputStream.Close()")
        ElseIf CanDownload Then
            code &= vbCrLf & "  ' save to an excel file"
            code &= String.Format(vbCrLf & "   workbook.Save(""{0}.xlsx"")", GetShortID())
        End If
        Return streamCode & code
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
        Return My.Resources.StringResource.ResourceManager.GetString(NameResKey, New Globalization.CultureInfo(culture))
    End Function

    Public Overridable Function GetDescriptionByCulture(culture As String) As String
        Return My.Resources.StringResource.ResourceManager.GetString(DescripResKey, New Globalization.CultureInfo(culture))
    End Function

    Protected Function AgentIsMac(userAgents() As String) As Boolean
        If userAgents.Length > 0 AndAlso userAgents(0).ToLower().Contains("macintosh") Then Return True
        Return False
    End Function

    Private Function ReadStreamToBase64(input As Stream) As String
        Using ms As New MemoryStream
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
            If _children Is Nothing Then _children = GetChildren()

            Return _children.ToArray()
        End Get
    End Property

    Private Function GetChildren() As List(Of ExampleBase)
        'INSTANT VB NOTE: The variable children was renamed since Visual Basic does not handle local variables named the same as class members well:
        Dim children_Renamed As New List(Of ExampleBase)
        Dim types() As Type = GetTypesRecursively(_namespace)
        Dim subNS As New HashSet(Of String)
        For Each type In types
            If type.Namespace = _namespace Then
                Dim child As ExampleBase = TryCast(Activator.CreateInstance(type), ExampleBase)
                If child.IsContainedInTree Then children_Renamed.Add(child)
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

        children_Renamed.Sort(New ExampleComparer)

        Return children_Renamed
    End Function

    'INSTANT VB NOTE: The variable id was renamed since Visual Basic does not handle local variables named the same as class members well:
    Public Function FindExample(id_Renamed As String) As ExampleBase
        Return FindExample(Me, id_Renamed)
    End Function

    'INSTANT VB NOTE: The variable id was renamed since Visual Basic does not handle local variables named the same as class members well:
    Private Function FindExample(example As ExampleBase, id_Renamed As String) As ExampleBase
        If example.ID = id_Renamed Then Return example

        Dim folderExample As FolderExample = TryCast(example, FolderExample)
        If folderExample IsNot Nothing Then
            For Each child In folderExample.Children
                Dim result As ExampleBase = FindExample(child, id_Renamed)
                If result IsNot Nothing Then Return result
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
                If item.IsUpdate OrElse item.IsNew Then Return True

                If IsUpdateRecursive(item) Then Return True
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
        _assembly = GetType(Examples).GetTypeInfo().Assembly
        _types = New List(Of Type)(_assembly.GetTypes())
        _types.Remove(_folderExampleType)
    End Sub

    Public Function GetTypesRecursively(ns As String) As Type()
        Return _types.FindAll(Function(type) type.Namespace IsNot Nothing AndAlso type.Namespace.StartsWith(ns) AndAlso type.GetTypeInfo().BaseType Is _exampleBaseType).ToArray()
    End Function
End Module

Module Program

End Module
