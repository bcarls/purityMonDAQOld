' This is taken from http://support.microsoft.com/kb/311340
' It's an attempt to replicate the ItemData in ComboBox functionality from VB6

Public Class MyList
    Private sName As String
    Private iID As Integer   'You can also declare this as String.

    Public Sub New()
        sName = ""
        iID = 0
    End Sub

    Public Sub New(ByVal Name As String, ByVal ID As Integer)
        sName = Name
        iID = ID
    End Sub

    Public Property Name() As String
        Get
            Return sName
        End Get

        Set(ByVal sValue As String)
            sName = sValue
        End Set
    End Property

    Public Property ItemData() As Integer
        Get
            Return iID
        End Get

        Set(ByVal iValue As Integer)
            iID = iValue
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return sName
    End Function
End Class
