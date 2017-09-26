
Public Class C_DEVICELIST


    Public Sub New()

    End Sub

    Public Sub New(Name As String, IP As String)
        Me.pName = Name
        Me.pIPAddress = IP
    End Sub

    Public Sub New(Name As String, IP As String, mTag As Object)
        Me.pName = Name
        Me.pIPAddress = IP
        Me.Tag = mTag
    End Sub

    Public pName As String
    Public pIPAddress As String
    Public pLogCount As Int32 = 0
    Public Tag As Object

End Class

Public Class C_DEVICELISTEXT : Inherits C_DEVICELIST

    Public pCompany As String
    Public pBranch As String

    Public Sub New()

    End Sub

    Public Sub New(Name As String, IP As String, Company As String, Branch As String)
        Me.pName = Name
        Me.pIPAddress = IP
        Me.pCompany = Company
        Me.pBranch = Branch
    End Sub
 
End Class