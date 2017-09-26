
Public NotInheritable Class C_EVENTHANDLERS

    Public Class AttHidNumArgs : Inherits EventArgs
        Public CardNumber As String
        Public AccessNo As String
    End Class

    Public Class NewUserEventArgs : Inherits EventArgs
        Public EnrollNumber As Integer
    End Class

    Public Class KeyPressEventArgs : Inherits EventArgs
        Public Key As Integer
    End Class

    Public Class VerifyEventsArgs : Inherits EventArgs
        Public UserID As Integer
    End Class

    Public Class AttTransactionEventArgs : Inherits EventArgs
        Public EnrollNumber As Integer
        Public IsInValid As Integer
        Public AttState As Integer
        Public VerifyMethod As Integer
        Public Year As Integer
        Public Month As Integer
        Public Day As Integer
        Public Hour As Integer
        Public Minute As Integer
        Public Second As Integer
        ' Public CardNo As String
    End Class

    Public Class AttTransactionEventArgsEx : Inherits EventArgs
        Public EnrollNumber As String
        Public IsInValid As Integer
        Public AttState As Integer
        Public VerifyMethod As Integer
        Public Year As Integer
        Public Month As Integer
        Public Day As Integer
        Public Hour As Integer
        Public Minute As Integer
        Public Second As Integer
        ' Public CardNo As String
    End Class

    Public Class HIdNumEventArgs : Inherits EventArgs
        Public CardNo As Integer
    End Class

    Public Class DeviceStatusEventArgs : Inherits EventArgs
        Public Property i As Int32
        Public Property max As Int32
    End Class

    Public Class FingerEnrollEventArgs : Inherits EventArgs
        Public Property EnrollNumber As Integer
        Public Property FingerIndex As Integer
        Public Property ActionResult As Integer
        Public Property TemplateLength As Integer
    End Class

    Public Class FingerEnrollExEventArgs : Inherits EventArgs
        Public Property EnrollNumber As String
        Public Property FingerIndex As Integer
        Public Property ActionResult As Integer
        Public Property TemplateLength As Integer
    End Class

End Class
