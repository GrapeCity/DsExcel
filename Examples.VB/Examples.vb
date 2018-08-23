Public Class Examples
    Private Shared _rootExample As FolderExample
    Shared Sub New()
        _rootExample = New RootExample(GetType(Examples).Namespace)
    End Sub
    Public Shared ReadOnly Property RootExample As FolderExample
        Get
            Return _rootExample
        End Get
    End Property
End Class
Public Class RootExample
    Inherits FolderExample
    Public Sub New(ns As String)
        MyBase.New(ns)
    End Sub
    Protected Overrides ReadOnly Property NameResKey As String
        Get
            Return "RootExample.Name"
        End Get
    End Property
    Protected Overrides ReadOnly Property DescripResKey As String
        Get
            Return "RootExample.Descrip"
        End Get
    End Property
End Class
