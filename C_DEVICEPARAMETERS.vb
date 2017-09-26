'Public Class C_ATTENDANCEPARAM
'    Implements IDisposable

'    Public IP As String
'    Public AccessNo As String
'    Public DtTime As DateTime
'    Public InOut As Integer
'    Public MinInterval As Integer
'    Public CardNo As String

'    Public Sub Dispose() Implements IDisposable.Dispose
'        MyBase.Finalize()
'    End Sub

'End Class

Public Class C_USERPARAM

    Public AccessNo As Object
    Public CardNo As String
    Public Name As String
    Public Password As String
    Public Priviledge As Integer
    Public Enabled As Boolean

End Class

Public Class C_USERPARAMEXT : Inherits C_USERPARAM

    'Public F1 As String
    'Public F2 As String
    'Public F3 As String
    'Public F4 As String
    'Public F5 As String
    'Public F6 As String
    'Public F7 As String
    'Public F8 As String
    'Public F9 As String
    'Public F10 As String

    Public Fingers As New Dictionary(Of Int32, String) 'because individual strings is too mainstream
    Public Faces As New Dictionary(Of Int32, String)

End Class
 

Public Structure StructNetOptions

    Public Const NetMask As String = "NetMask"
    Public Const GATEIPAddress As String = "GATEIPAddress"

End Structure