Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports zkemkeeper
Imports z.Device.C_EVENTHANDLERS
Imports System.Threading.Tasks
Imports System.Security.Permissions

''' <summary>
'''  LJ 20130309
''' </summary>
''' <remarks></remarks>
Public Class C_DEVICE
    Implements IDisposable

#Region " CONSTRUCTORS "

    Public Sub New()
        Try
            AddHandler AppDomain.CurrentDomain.UnhandledException, New UnhandledExceptionEventHandler(AddressOf unhndled)
            Me.mZ = New CZKEM
        Catch ex As AccessViolationException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub New(ID As Int32)
        Try
            AddHandler AppDomain.CurrentDomain.UnhandledException, New UnhandledExceptionEventHandler(AddressOf unhndled)
            Me.mZ = New CZKEM
            Me.mID = ID
        Catch ex As AccessViolationException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub New(ID As Int32, IP As String, Port As Int32)
        Try
            AddHandler AppDomain.CurrentDomain.UnhandledException, New UnhandledExceptionEventHandler(AddressOf unhndled)
            Me.mZ = New CZKEM
            Me.mIP = IP
            Me.mPort = Port
            Me.mID = ID
        Catch ex As AccessViolationException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub New(ID As Int32, ComPort As Int32, Optional BaudRate As Int32 = 115200)
        Try
            AddHandler AppDomain.CurrentDomain.UnhandledException, New UnhandledExceptionEventHandler(AddressOf unhndled)
            Me.mZ = New CZKEM
            Me.mID = ID
            Me.mComPort = ComPort
            Me.mBaudRate = BaudRate
        Catch ex As AccessViolationException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' LJ 20130912
    ''' Dont use Disposable if realtime is invoke
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            If Me.mConnected = True Then
                Me.Disconnect()
            End If

            Me.mPartialLogs = Nothing
            Me.mFullLogs = Nothing
            Me.mZ = Nothing

        Catch ex As AccessViolationException
            Throw ex
        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.SuppressFinalize(Me)
        End Try
    End Sub

#End Region

#Region " VARIABLES "

    Private WithEvents mZ As CZKEM

    Private mID As Int32
    Private mIP As String
    Private mPort As Integer

    '-- Com
    Private mComPort As Integer
    Private mBaudRate As Integer = 115200

    Private mMachineNumber As Integer = 1
    Private mConnected As Boolean = False
    Private mDeviceName As String = ""

    Private mCurrenFingerIndex As Int32 = 0

#End Region

#Region " MUST INHERIT EVENTS "

    Public Event hConnected As dgOnConnected
    Public Event hDisConnected As dgOnDisconnected
    Public Event hFinger As dgOnFinger
    Public Event hVerify As dgOnVerify
    Public Event hTransaction As dgOnAttransaction
    Public Event hTransactionEx As dgOnAttransactionEx
    Public Event hKeyPress As dgOnKeyPress
    Public Event hHidNum As dgOnHidNum
    Public Event hFingerEnroll As dgOnFingerEnroll
    Public Event hFingerEnrollEx As dgOnFingerEnrollEx

#End Region

#Region " EVENTS "

    Private Sub unhndled(Sender As Object, e As UnhandledExceptionEventArgs)
        IO.File.AppendAllText(IO.Path.Combine(Environment.CurrentDirectory, "ScnProc.InSysLog"), Date.Now & "-> " & e.ExceptionObject.GetHashCode & vbCrLf)
    End Sub

    Private Sub OnFinger()
        RaiseEvent hFinger(Me, EventArgs.Empty)
    End Sub

    Private Sub OnVerify(UserID As Integer)
        Dim e As New VerifyEventsArgs
        e.UserID = UserID
        RaiseEvent hVerify(Me, e)
    End Sub

    <MTAThread> _
    Private Sub OnAttTransaction(ByVal EnrollNumber As Integer,
                                 ByVal IsInValid As Integer,
                                 ByVal AttState As Integer,
                                 ByVal VerifyMethod As Integer,
                                 ByVal Year As Integer,
                                 ByVal Month As Integer,
                                 ByVal Day As Integer,
                                 ByVal Hour As Integer,
                                 ByVal Minute As Integer,
                                 ByVal Second As Integer)

        Dim e As New AttTransactionEventArgs
        e.EnrollNumber = EnrollNumber
        'e.CardNo = getCardNo(CStr(EnrollNumber))
        e.IsInValid = IsInValid
        e.AttState = AttState
        e.VerifyMethod = VerifyMethod
        e.Year = Year
        e.Month = Month
        e.Day = Day
        e.Hour = Hour
        e.Minute = Minute
        e.Second = Second

        Dim p As New dgRTransaction(AddressOf RTransaction)
        Dim thrd As New Threading.Thread(Sub() p.Invoke(e))
        thrd.Start()

    End Sub

    Private Sub OnKeyPress(Key As Integer)
        Dim e As New KeyPressEventArgs
        e.Key = Key
        RaiseEvent hKeyPress(Me, e)
    End Sub

    <MTAThread> _
    Private Sub OnAttTransactionEx(ByVal EnrollNumber As String,
                                   ByVal IsInValid As Integer,
                                   ByVal AttState As Integer,
                                   ByVal VerifyMethod As Integer,
                                   ByVal Year As Integer,
                                   ByVal Month As Integer,
                                   ByVal Day As Integer,
                                   ByVal Hour As Integer,
                                   ByVal Minute As Integer,
                                   ByVal Second As Integer,
                                   ByVal WorkCode As Integer)
        Dim e As New AttTransactionEventArgsEx
        e.EnrollNumber = EnrollNumber
        'e.CardNo = getCardNo(CStr(EnrollNumber)) // CardNo not part of  transaction
        e.IsInValid = IsInValid
        e.AttState = AttState
        e.VerifyMethod = VerifyMethod
        e.Year = Year
        e.Month = Month
        e.Day = Day
        e.Hour = Hour
        e.Minute = Minute
        e.Second = Second

        Dim p As New dgRTransactionEx(AddressOf RTransactionEx)
        Dim thrd As New Threading.Thread(Sub() p.Invoke(e))
        thrd.Start()

    End Sub

    Private Sub OnConnected()
        Me.mConnected = True
        RaiseEvent hConnected(Me, EventArgs.Empty)
    End Sub

    Private Sub OnDisConnected()
        Me.mConnected = False
        RaiseEvent hDisConnected(Me, EventArgs.Empty)
    End Sub

    Private Sub OnHidnum(CardNumber As Integer) Handles mZ.OnHIDNum
        Dim e As New HIdNumEventArgs
        e.CardNo = CardNumber
        RaiseEvent hHidNum(Me, e)
    End Sub

    Private Sub OnEnrollFinger(EnrollNumber As Integer, FingerIndex As Integer, ActionResult As Integer, TemplateLenght As Integer)
        Dim e As New FingerEnrollEventArgs
        e.EnrollNumber = EnrollNumber
        e.FingerIndex = Me.mCurrenFingerIndex 'FingerIndex
        e.ActionResult = ActionResult
        e.TemplateLength = TemplateLenght
        RaiseEvent hFingerEnroll(Me, e)
    End Sub

    Private Sub OnEnrollFingerEx(EnrollNumber As String, FingerIndex As Integer, ActionResult As Integer, TemplateLenght As Integer)
        Dim e As New FingerEnrollExEventArgs
        e.EnrollNumber = EnrollNumber
        e.FingerIndex = Me.mCurrenFingerIndex 'FingerIndex
        e.ActionResult = ActionResult
        e.TemplateLength = TemplateLenght
        RaiseEvent hFingerEnrollEx(Me, e)
    End Sub

    Public Event ShowStatus As dgShowStatus


#End Region

#Region " DELEGATES "

    Public Delegate Sub dgOnConnected(Sender As C_DEVICE, e As EventArgs)
    Public Delegate Sub dgOnDisconnected(Sender As C_DEVICE, e As EventArgs)
    Public Delegate Sub dgOnFinger(Sender As C_DEVICE, e As EventArgs)
    Public Delegate Sub dgOnVerify(Sender As C_DEVICE, e As VerifyEventsArgs)
    Public Delegate Sub dgOnAttransaction(Sender As C_DEVICE, e As AttTransactionEventArgs)
    Public Delegate Sub dgOnAttransactionEx(Sender As C_DEVICE, e As AttTransactionEventArgsEx)
    Public Delegate Sub dgOnKeyPress(Sender As C_DEVICE, e As KeyPressEventArgs)
    Public Delegate Sub dgOnHidNum(Sender As C_DEVICE, e As HIdNumEventArgs)

    Private Delegate Sub dgRTransactionEx(e As AttTransactionEventArgsEx)
    Private Delegate Sub dgRTransaction(e As AttTransactionEventArgs)

    Public Delegate Sub dgOnFingerEnroll(Sender As C_DEVICE, e As FingerEnrollEventArgs)
    Public Delegate Sub dgOnFingerEnrollEx(Sender As C_DEVICE, e As FingerEnrollExEventArgs)

    Public Delegate Sub dgShowStatus(Sender As C_DEVICE, e As DeviceStatusEventArgs)

    Public Delegate Function dgConnect(ByVal ConnectType As eConnectType) As Boolean

#End Region

#Region " METHODS "

    ''' <summary>
    ''' This Requires Admin Priviledge
    ''' </summary>
    ''' <param name="zkempath"></param>
    ''' <param name="silent"></param>
    ''' <remarks></remarks>
    <STAThread> _
    Public Shared Sub Register(zkempath As String, Optional silent As Boolean = True) '<PrincipalPermission(SecurityAction.Demand)> _
        Try
            Dim cmd As New Process
            With cmd
                .StartInfo.FileName = "cmd.exe"
                .StartInfo.RedirectStandardInput = True
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.CreateNoWindow = True
                .StartInfo.UseShellExecute = False
                .Start()
                .StandardInput.WriteLine("@echo off")
                .StandardInput.WriteLine(String.Format("cd {0}", zkempath))
                .StandardInput.WriteLine(String.Format("{0}:", zkempath.Substring(0, 1)))
                .StandardInput.WriteLine("regsvr32 zkemkeeper.dll {0}", IIf(silent, "/s", ""))
                .StandardInput.Flush()
                .StandardInput.Close()
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    <MTAThread> _
    Public Sub Disconnect()
        Try
            If Me.mConnected Then
                mZ.Disconnect()
            End If
            Me.mConnected = False
            RaiseEvent hDisConnected(Me, EventArgs.Empty)
        Catch ex As AccessViolationException
            'Throw ex Suppress
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub Beep()
        mZ.Beep(0)
    End Sub

    Public Function SetToEnglish() As Boolean
        SetToEnglish = mZ.SetDeviceInfo(mMachineNumber, 3, 0)

    End Function

    Public Function GetLanguage() As Integer
        Dim x As Integer
        mZ.GetDeviceInfo(mMachineNumber, 3, x)
        Return x
    End Function

    Public Function LoadToMemory() As Boolean
        Return mZ.BatchUpdate(mMachineNumber)
    End Function

    Public Sub RefreshKey()
        For i As Int32 = 1 To 6 Step 1
            mZ.SetCustomizeAttState(mMachineNumber, i, i)
        Next
    End Sub

    Public Function SetVoice(index As Int32, file As String) As Boolean
        'For i As Int32 = 0 To 11 Step 1
        '    mZ.EnableCustomizeVoice(mMachineNumber, 0, 1)
        'Next
        Return mZ.SetCustomizeVoice(mMachineNumber, index, file)
    End Function

    Public Sub GetVoice()
        If mZ.GetDataFile(mMachineNumber, 0, "E_1.wav") Then
            mZ.SaveTheDataToFile(mMachineNumber, "c:\E_1.wav", 0)
        End If
    End Sub

    Public Sub RefreshEvent()
        Try
            mZ.RefreshData(mMachineNumber)
        Catch
        End Try
    End Sub

    Public Sub Listen()
        AddHandler mZ.OnFinger, AddressOf OnFinger
        AddHandler mZ.OnVerify, AddressOf OnVerify
        AddHandler mZ.OnAttTransaction, AddressOf OnAttTransaction
        AddHandler mZ.OnKeyPress, AddressOf OnKeyPress
        AddHandler mZ.OnAttTransactionEx, AddressOf OnAttTransactionEx
        AddHandler mZ.OnConnected, AddressOf OnConnected
        AddHandler mZ.OnDisConnected, AddressOf OnDisConnected
        AddHandler mZ.OnEnrollFinger, AddressOf OnEnrollFinger
        AddHandler mZ.OnEnrollFingerEx, AddressOf OnEnrollFingerEx
    End Sub

    Public Sub UnListen()
        RemoveHandler mZ.OnKeyPress, AddressOf OnKeyPress
        RemoveHandler mZ.OnFinger, AddressOf OnFinger
        RemoveHandler mZ.OnVerify, AddressOf OnVerify
        RemoveHandler mZ.OnAttTransaction, AddressOf OnAttTransaction
        RemoveHandler mZ.OnAttTransactionEx, AddressOf OnAttTransactionEx
        RemoveHandler mZ.OnConnected, AddressOf OnConnected
        RemoveHandler mZ.OnDisConnected, AddressOf OnDisConnected
        RemoveHandler mZ.OnEnrollFinger, AddressOf OnEnrollFinger
        RemoveHandler mZ.OnEnrollFingerEx, AddressOf OnEnrollFingerEx

    End Sub

    Public Function setIpAddress(ByVal IpAddress As String) As Boolean
        Dim valid As Boolean = False
        valid = mZ.SetDeviceIP(1, IpAddress)
        Return valid
    End Function

    Public Function setNetOptions(ByVal StructNetOptionValue As eNetOptions, ByVal val As String) As Boolean
        setNetOptions = setSysOption(StructNetOptionValue.ToString, val)
    End Function
    '--

    Public Function setSysOption(param As String, val As String) As Boolean
        Dim vl As Boolean = False
        vl = mZ.SetSysOption(mMachineNumber, param, val)
        Return vl
    End Function

    Public Delegate Function dgSSR_GetTimeLog(MinIntereval As Integer, DayInterval As Integer, GetFullAttLogs As Boolean) As String()

    <MTAThread> _
    Public Function SSR_GetTimeLog(MinInterval As Integer, DayInterval As Integer, GetFullAttLogs As Boolean) As String()
        Try

            Dim dwEnrollNumber As String = ""
            Dim dwVerifyMode As Integer
            Dim dwInOutMode As Integer
            Dim dwYear As Integer
            Dim dwMonth As Integer
            Dim dwDay As Integer
            Dim dwHour As Integer
            Dim dwMinute As Integer
            Dim dwSecond As Integer
            Dim dwWorkCode As Integer
            Dim AttList() As String = New String() {}

            Try
                SyncLock mZ

                    Dim stat As New DeviceStatusEventArgs
                    stat.max = CInt(getRecordCount(C_DEVICE.eRecordCount.Attendance_records))
                    Dim b As Boolean = False

                    b = mZ.ReadAllGLogData(mMachineNumber)
                    If b Then
                        While mZ.SSR_GetGeneralLogData(mMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkCode)

                            SyncLock AttList

                                Dim inout As Int32 = dwInOutMode
                                'If dwInOutMode <> 0 And dwInOutMode <> 1 Then
                                '    inout = 2
                                'Else
                                '    inout = dwInOutMode + 1
                                'End If

                                Dim dttime As DateTime = CDate("1/1/1990")
                                Try
                                    dttime = CDate(dwMonth.ToString & "/" & dwDay & "/" & dwYear.ToString & " " & dwHour.ToString & ":" & dwMinute.ToString)
                                Catch ex As Exception
                                    dttime = CDate("1/1/1990")
                                End Try

                                stat.i += 1

                                'If CInt(stat.i Mod 100) = 0 Then
                                'RaiseEvent ShowStatus(Me, stat)
                                'End If

                                If GetFullAttLogs = False Then
                                    Dim DayCheck As Date = DateAdd(DateInterval.Day, DayInterval, Date.Now)
                                    If CDate(Format(dttime, "MM/dd/yyyy")) < CDate(Format(DayCheck, "MM/dd/yyyy")) Then
                                        Continue While
                                    End If
                                End If

                                Array.Resize(AttList, AttList.Length + 1)
                                AttList(AttList.Length - 1) = mIP & vbTab & dwEnrollNumber & vbTab & dttime & vbTab & inout & vbTab & MinInterval



                                ' Threading.Thread.Sleep(200)

                            End SyncLock

                        End While
                    Else
                        Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                    End If

                End SyncLock

                If AttList.Length = 0 Then
                    Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                End If

                Return AttList
            Catch ex As AccessViolationException
                Throw ex
            Catch ex As OutOfMemoryException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                AttList = Nothing
                Me.RefreshEvent()
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Delegate Function dgGetTimeLog(MinIntereval As Integer, DayInterval As Integer, GetFullAttLogs As Boolean) As String()

    <MTAThread> _
    Public Function GetTimeLog(MinInterval As Integer, DayInterval As Integer, GetFullAttLogs As Boolean) As String()
        Try
            SyncLock mZ
                Dim dwTMachineNumber As Integer
                Dim dwEnrollNumber As Integer
                Dim dwEmachineNumber As Integer
                Dim dwVerifyMode As Integer
                Dim dwInOutMode As Integer
                Dim dwYear As Integer
                Dim dwMonth As Integer
                Dim dwDay As Integer
                Dim dwHour As Integer
                Dim dwMinute As Integer

                Dim AttList() As String = New String() {}
                Dim stat As DeviceStatusEventArgs

                Try
                    SyncLock Me

                        stat = New DeviceStatusEventArgs

                        stat.max = CInt(getRecordCount(C_DEVICE.eRecordCount.Attendance_records))
                        stat.i = 0
                        RaiseEvent ShowStatus(Me, stat)

                        Dim b As Boolean = False

                        b = mZ.ReadAllGLogData(mMachineNumber)
                        If b Then

                            While mZ.GetGeneralLogData(mMachineNumber, dwTMachineNumber, dwEnrollNumber, dwEmachineNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute)

                                SyncLock AttList

                                    Dim inout As Int32 = dwInOutMode
                                    'If dwInOutMode <> 0 And dwInOutMode <> 1 Then
                                    '    inout = 2
                                    'Else
                                    '    inout = dwInOutMode + 1
                                    'End If

                                    Dim dttime As DateTime = CDate("1/1/1990")
                                    Try
                                        dttime = CDate(dwMonth.ToString & "/" & dwDay & "/" & dwYear.ToString & " " & dwHour.ToString & ":" & dwMinute.ToString)
                                    Catch ex As Exception
                                        dttime = CDate("1/1/1990")
                                    End Try

                                    stat.i += 1

                                    '                                   If CInt(stat.i Mod 100) = 0 Then
                                    'RaiseEvent ShowStatus(Me, stat)
                                    '                                    End If

                                    If GetFullAttLogs = False Then
                                        Dim DayCheck As Date = DateAdd(DateInterval.Day, DayInterval, Date.Now)
                                        If CDate(Format(dttime, "MM/dd/yyyy")) < CDate(Format(DayCheck, "MM/dd/yyyy")) Then
                                            Continue While
                                        End If
                                    End If

                                    Array.Resize(AttList, AttList.Length + 1)
                                    AttList(AttList.Length - 1) = mIP & vbTab & dwEnrollNumber & vbTab & dttime & vbTab & inout & vbTab & MinInterval


                                    ' Threading.Thread.Sleep(200)

                                End SyncLock

                            End While
                        Else
                            Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                        End If

                    End SyncLock

                    If AttList.Length = 0 Then
                        Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                    End If

                    Return AttList
                Catch ex As AccessViolationException
                    Throw ex
                Catch m As OutOfMemoryException
                    Throw m
                Catch ex As Exception
                    Throw ex
                Finally
                    AttList = Nothing
                    Me.RefreshEvent()
                End Try

            End SyncLock
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Delegate Sub dgVoid_GetTimeLog(MinIntereval As Integer, DayInterval As Integer)

    ''' <summary>
    ''' LJ 20130912
    ''' Get Full and Partial Logs for Algo 9 devices
    ''' </summary>
    ''' <param name="MinInterval"></param>
    ''' <param name="DayInterval"></param>
    ''' <remarks>
    ''' 
    ''' Get Full Logs in pFullLogs Property
    ''' Get Partial Logs in pPartial Logs Property
    ''' 
    ''' </remarks>
    <MTAThread> _
    Public Sub GetTimeLog(MinInterval As Integer, DayInterval As Integer, Optional action As Action(Of Int32) = Nothing, Optional log As Action(Of String) = Nothing)
        Try
            SyncLock mZ
                Dim dwTMachineNumber As Integer
                Dim dwEnrollNumber As Integer
                Dim dwEmachineNumber As Integer
                Dim dwVerifyMode As Integer
                Dim dwInOutMode As Integer
                Dim dwYear As Integer
                Dim dwMonth As Integer
                Dim dwDay As Integer
                Dim dwHour As Integer
                Dim dwMinute As Integer

                Me.mFullLogs = New String() {}
                Me.mPartialLogs = New String() {}

                'Dim stat As DeviceStatusEventArgs
                Dim DayCheck As Date

                Try
                    SyncLock Me

                        'stat = New DeviceStatusEventArgs

                        'stat.max = CInt(getRecordCount(C_DEVICE.eRecordCount.Attendance_records))
                        'stat.i = 0
                        'RaiseEvent ShowStatus(Me, stat)

                        Dim i As Int32 = 0
                        Dim b As Boolean = False

                        b = mZ.ReadAllGLogData(mMachineNumber)
                        If b Then

                            While mZ.GetGeneralLogData(mMachineNumber, dwTMachineNumber, dwEnrollNumber, dwEmachineNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute)

                                SyncLock Me.mPartialLogs

                                    'Dim inout As Int32
                                    'If dwInOutMode <> 0 And dwInOutMode <> 1 Then
                                    '    inout = 2
                                    'Else
                                    '    inout = dwInOutMode + 1
                                    'End If

                                    Dim dttime As DateTime = CDate("1/1/1990")
                                    Try
                                        dttime = New Date(dwYear, dwMonth, dwDay, dwHour, dwMinute, 0) 'CDate(dwMonth.ToString & "/" & dwDay & "/" & dwYear.ToString & " " & dwHour.ToString & ":" & dwMinute.ToString)
                                    Catch ex As Exception
                                        dttime = CDate("1/1/1990")
                                    End Try

                                    i += 1

                                    If action IsNot Nothing Then action(i)

                                    'Get Full Logs
                                    Array.Resize(Me.mFullLogs, Me.mFullLogs.Length + 1)
                                    Me.mFullLogs(Me.mFullLogs.Length - 1) = mIP & vbTab & dwEnrollNumber & vbTab & dttime & vbTab & dwInOutMode & vbTab & dwVerifyMode & vbTab & MinInterval

                                    DayCheck = DateAdd(DateInterval.Day, DayInterval, Date.Now)
                                    If CDate(Format(dttime, "MM/dd/yyyy")) < CDate(Format(DayCheck, "MM/dd/yyyy")) Then
                                        Continue While
                                    End If

                                    Array.Resize(Me.mPartialLogs, Me.mPartialLogs.Length + 1)
                                    Me.mPartialLogs(Me.mPartialLogs.Length - 1) = mIP & vbTab & dwEnrollNumber & vbTab & dttime & vbTab & dwInOutMode & vbTab & dwVerifyMode & vbTab & MinInterval
                                    If log IsNot Nothing Then log(Me.mPartialLogs(Me.mPartialLogs.Length - 1))

                                End SyncLock

                            End While
                        Else
                            Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                        End If

                    End SyncLock

                    If mFullLogs.Length = 0 Then
                        Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                    End If

                Catch ex As AccessViolationException
                    Throw ex
                Catch m As OutOfMemoryException
                    Throw m
                Catch ex As Exception
                    Throw ex
                Finally
                    Me.RefreshEvent()
                End Try

            End SyncLock
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Delegate Sub dgVoid_SSR_GetTimeLog(MinIntereval As Integer, DayInterval As Integer)

    ''' <summary>
    ''' LJ 20130912
    ''' Get Full and Partial Logs for Algo 10 devices
    ''' </summary>
    ''' <param name="MinInterval"></param>
    ''' <param name="DayInterval"></param>
    ''' <remarks>
    ''' 
    ''' Get Full Logs in pFullLogs Property
    ''' Get Partial Logs in pPartial Logs Property
    ''' 
    ''' </remarks>
    <MTAThread> _
    Public Sub SSR_GetTimeLog(MinInterval As Integer, DayInterval As Integer, Optional action As Action(Of Int32) = Nothing, Optional log As Action(Of String) = Nothing)
        Try

            Dim dwEnrollNumber As String = ""
            Dim dwVerifyMode As Integer
            Dim dwInOutMode As Integer
            Dim dwYear As Integer
            Dim dwMonth As Integer
            Dim dwDay As Integer
            Dim dwHour As Integer
            Dim dwMinute As Integer
            Dim dwSecond As Integer
            Dim dwWorkCode As Integer

            Me.mFullLogs = New String() {}
            Me.mPartialLogs = New String() {}

            Dim DayCheck As Date

            Try
                SyncLock mZ

                    'Dim stat As New DeviceStatusEventArgs
                    'stat.max = CInt(getRecordCount(C_DEVICE.eRecordCount.Attendance_records))
                    Dim i As Int32 = 0
                    Dim b As Boolean = False

                    b = mZ.ReadAllGLogData(mMachineNumber)
                    If b Then
                        While mZ.SSR_GetGeneralLogData(mMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkCode)

                            SyncLock Me.mFullLogs

                                'Dim inout As Int32
                                'If dwInOutMode <> 0 And dwInOutMode <> 1 Then
                                '    inout = 2
                                'Else
                                '    inout = dwInOutMode + 1
                                'End If

                                Dim dttime As DateTime = CDate("1/1/1990")
                                Try
                                    dttime = New Date(dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond) 'CDate(dwMonth.ToString & "/" & dwDay & "/" & dwYear.ToString & " " & dwHour.ToString & ":" & dwMinute.ToString)
                                Catch ex As Exception
                                    dttime = CDate("1/1/1990")
                                End Try

                                i += 1

                                'If CInt(stat.i Mod 100) = 0 Then
                                'RaiseEvent ShowStatus(Me, stat)
                                'End If
                                If action IsNot Nothing Then action(i)

                                'Get Full Logs
                                Array.Resize(Me.mFullLogs, Me.mFullLogs.Length + 1)
                                Me.mFullLogs(Me.mFullLogs.Length - 1) = mIP & vbTab & dwEnrollNumber & vbTab & dttime & vbTab & dwInOutMode & vbTab & dwVerifyMode & vbTab & MinInterval

                                DayCheck = DateAdd(DateInterval.Day, DayInterval, Date.Now)
                                If CDate(Format(dttime, "MM/dd/yyyy")) < CDate(Format(DayCheck, "MM/dd/yyyy")) Then
                                    Continue While
                                End If

                                Array.Resize(Me.mPartialLogs, Me.mPartialLogs.Length + 1)
                                Me.mPartialLogs(Me.mPartialLogs.Length - 1) = mIP & vbTab & dwEnrollNumber & vbTab & dttime & vbTab & dwInOutMode & vbTab & dwVerifyMode & vbTab & MinInterval
                                If log IsNot Nothing Then log(Me.mPartialLogs(Me.mPartialLogs.Length - 1))

                            End SyncLock

                        End While
                    Else
                        Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                    End If

                End SyncLock

                If Me.mFullLogs.Length = 0 Then
                    Throw New Exception("Could not Get Logs from Device IP:" & Me.mIP)
                End If

            Catch ex As AccessViolationException
                Throw ex
            Catch ex As OutOfMemoryException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                Me.RefreshEvent()
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    <MTAThread> _
    Private Sub RTransaction(e As AttTransactionEventArgs)
        Try
            RaiseEvent hTransaction(Me, e)
        Catch ex As Exception
        Finally
            Threading.Thread.CurrentThread.Abort()
        End Try
    End Sub

    <MTAThread> _
    Private Sub RTransactionEx(e As AttTransactionEventArgsEx)
        Try
            RaiseEvent hTransactionEx(Me, e)
        Catch ex As Exception
        Finally
            Threading.Thread.CurrentThread.Abort()
        End Try
    End Sub

    Public Sub CancelOperation()
        mZ.CancelOperation()
    End Sub

    Public Function StartFingerEnroll(AccessNo As String, Finger As Int32, Platform As String) As Boolean

        Dim p As Boolean = False
        mZ.CancelOperation()
        mZ.SSR_DelUserTmpExt(mMachineNumber, AccessNo, Finger)

        If Platform.Contains("TFT") Then
            If mZ.StartEnrollEx(AccessNo, Finger, 3) = True Then
                p = True
                Me.mCurrenFingerIndex = Finger
            Else
                p = False
            End If
        Else
            If mZ.StartEnroll(CInt(AccessNo), Finger) Then
                p = True
                Me.mCurrenFingerIndex = Finger
            Else
                p = False
            End If
        End If

        Return p
    End Function

    Public Function DeleteFingerDataIndex(AccessNo As Int32, Finger As Int32) As Boolean
        Return mZ.DelUserTmp(mMachineNumber, AccessNo, Finger)
    End Function

    Public Function SSR_DeleteFingerDataIndex(AccessNo As String, Finger As Int32) As Boolean
        Return mZ.SSR_DelUserTmp(mMachineNumber, AccessNo, Finger)
    End Function

    Public Sub StartFingerEnrollIdentify()
        mZ.StartIdentify()
        mZ.RefreshData(mMachineNumber)
    End Sub

    Public Function SaveCardNo(AccessNo As Integer, CardNo As String) As Boolean
        Try
            Dim b As Boolean = False
            b = mZ.SetStrCardNumber(CardNo)
            b = mZ.SetUserInfo(mMachineNumber, AccessNo, "", "", 0, True)
            Return b
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function EnableDevice(enbld As Boolean) As Boolean
        Try
            Return Me.mZ.EnableDevice(mMachineNumber, enbld)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ClearAll() As Boolean
        Me.mZ.ClearData(mMachineNumber, eClearFlag.User_Information)
        Me.mZ.ClearGLog(mMachineNumber)
        Return True
    End Function

    Public Function Clear(flag As eClearFlag) As Boolean
        Return Me.mZ.ClearData(mMachineNumber, flag)
    End Function

    Public Function BatchUpdate() As Boolean
        Return Me.mZ.BatchUpdate(mMachineNumber)
    End Function

    Public Function SupportsCard() As Boolean
        SupportsCard = False
        Dim x As Integer
        mZ.GetCardFun(mMachineNumber, x)
        If x <> 0 Then
            SupportsCard = True
        End If
    End Function

    ''' <summary>
    ''' Push the events to wait loop
    ''' this must be use to console app only, to raise events
    ''' </summary>
    ''' <remarks>
    ''' LJ 20140513
    ''' </remarks>
    Public Sub ConsoleActivateEvents()
        System.Windows.Forms.Application.Run() ' the fuck! need to call to raise events
    End Sub

    Public Sub ClearMemoryLogs()
        Me.mPartialLogs = Nothing
        Me.mFullLogs = Nothing
    End Sub

#End Region

#Region " FUNCTIONS "

    <MTAThread> _
    Public Function Connect() As Boolean
        Dim b As Boolean = False
        Try
            Me.Connect(eConnectType.Net)
            b = True
        Catch ex As Exception
            b = False
        End Try
        Return b
    End Function

    ''' <summary>
    ''' On Connection returns, please see events, hconnected, hdisconneted
    ''' </summary>
    ''' <param name="ConnectType"></param>
    ''' <remarks></remarks>
    <MTAThread> _
    Public Sub Connect(Optional ByVal ConnectType As eConnectType = eConnectType.Net) 'As Boolean
        Try
            'SyncLock Me.mZ
            '    Select Case ConnectType
            '        Case eConnectType.Net
            '            mConnected = mZ.Connect_Net(mIP, mPort)
            '        Case eConnectType.Com : mConnected = mZ.Connect_Com(mComPort, mMachineNumber, mBaudRate)
            '        Case eConnectType.RS232 : mConnected = mZ.Connect_USB(mMachineNumber)
            '    End Select

            '    If mConnected = True Then
            '        Me.RegEvents(mMachineNumber, 65535)
            '        RaiseEvent hConnected(Me, EventArgs.Empty)
            '    Else
            '        RaiseEvent hDisConnected(Me, EventArgs.Empty)
            '    End If

            '    Return mConnected
            'End SyncLock

            Using conn As New Connectify
                AddHandler conn.Connect, AddressOf OnConnected
                AddHandler conn.DisConnect, AddressOf OnDisConnected

                'Dim thrd As New Threading.Thread(Sub() conn.StartConnect(Me, ConnectType))
                'thrd.Start()
                conn.StartConnect(Me, ConnectType)
            End Using
        Catch ex As AccessViolationException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function ClearAdministrators() As Boolean
        Try
            ClearAdministrators = mZ.ClearAdministrators(mMachineNumber)
        Catch ex As Exception
            ClearAdministrators = False
        End Try
        Return ClearAdministrators
    End Function

    Function GetMAC() As String
        Dim MAC As String = ""
        mZ.GetDeviceMAC(mMachineNumber, MAC)
        Return MAC
    End Function

    Function CanPing() As Boolean
        Dim tryCount = 0

        Using png As New Net.NetworkInformation.Ping
            Dim pngReply As Net.NetworkInformation.PingReply = png.Send(Me.mIP, 10000)

            While pngReply.Status <> Net.NetworkInformation.IPStatus.Success

                If tryCount >= 1 Then
                    'Throw New Exception("Cannot Ping Device")
                    Return False
                End If

                pngReply = png.Send(Me.mIP, 10000)

                Threading.Thread.Sleep(1000)
                tryCount += 1
            End While
        End Using

        Return True
    End Function

    Private Function GetSerialNumber() As String
        Dim s As String = ""
        mZ.GetSerialNumber(mMachineNumber, s)
        Return s
    End Function

    Public Function GetPlatform() As String
        Dim strPC As String = ""
        mZ.GetPlatform(mMachineNumber, strPC)
        Return strPC
    End Function

    Public Function RestartDevice() As Boolean
        Try
            Return mZ.RestartDevice(mMachineNumber)
        Catch EX As Exception
            Throw EX
        End Try
    End Function

    Public Function getCardNo(ByVal EnrollNumber As String) As String
        Dim CrdNo As String = ""
        Try
            'Dim lol As Integer = 0
            'Dim lolb As Byte
            'mZ.GetUserInfoEx(mMachineNumber, CInt(EnrollNumber), lol, lolb)
            mZ.GetStrCardNumber(CrdNo)
        Catch ex As Exception
        End Try

        Return CrdNo
    End Function

    Public Function GetTime() As DateTime
        Dim idwYear As Integer
        Dim idwMonth As Integer
        Dim idwDay As Integer
        Dim idwHour As Integer
        Dim idwMinute As Integer
        Dim idwSecond As Integer
        mZ.GetDeviceTime(mMachineNumber, idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond)

        Return New DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond)
    End Function

    Public Function SetTime(ByVal pdateTime As DateTime) As Boolean
        ' If mZ.SetDeviceTime(mMachineNumber) Then
        Return mZ.SetDeviceTime2(mMachineNumber, CInt(Format(pdateTime, "yyyy")), CInt(Format(pdateTime, "MM")), CInt(Format(pdateTime, "dd")), CInt(Format(pdateTime, "HH")), CInt(Format(pdateTime, "mm")), CInt(Format(pdateTime, "ss")))
        'Else
        'Return False
        'End If
    End Function

    Public Function RegEvents(ByVal reMachineNumber As Integer, ByVal reEventMask As Integer) As Boolean
        Return mZ.RegEvent(reMachineNumber, reEventMask)
    End Function

    Public Function getRecordCount(ByVal intStatus As eRecordCount) As Integer
        Try
            Dim i As Integer = 0
            Dim pd As Boolean = mZ.GetDeviceStatus(mMachineNumber, intStatus, i)
            Return i
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ClearLogs() As Boolean
        Try
            Return mZ.ClearGLog(mMachineNumber)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SetUserInfoEx(AccesNo As Int32, verstyle As eVerifyStyle) As Boolean
        Dim b As Boolean = False
        Dim bt As Byte
        b = mZ.SetUserInfoEx(mMachineNumber, AccesNo, verstyle, bt)
        Return b
    End Function

    Public Function SetUserInfoEx(AccesNo As Int32, verstyle As eVerifyStyle_TFT) As Boolean
        Dim b As Boolean = False
        Dim bt As Byte
        b = mZ.SetUserInfoEx(mMachineNumber, AccesNo, verstyle, bt)
        Return b
    End Function

    'Public Function SetUserInfoEx(AccesNo As String, verstyle As eVerifyStyle_TFT) As Boolean
    '    Dim b As Boolean = False
    '    Dim bt As Byte
    '    b = mZ.SetUserInfoEx().SetUserInfoEx(mMachineNumber, AccesNo, verstyle, bt)
    '    Return b
    'End Function

    Public Function GetUserInfo() As List(Of C_USERPARAM)
        Try
            Dim cList As New List(Of C_USERPARAM)

            Dim b As Boolean
            Dim pPrivelege As Integer
            Dim pName As String = ""
            Dim pPassword As String = ""
            Dim pEnabled As Boolean
            Dim dwEnrollNumber As Integer
            Dim dwMachinePrivilege As Integer

            Dim name As String = ""

            b = mZ.ReadAllUserID(mMachineNumber)
            If b Then
                While mZ.GetAllUserInfo(mMachineNumber, dwEnrollNumber, pName, pPassword, dwMachinePrivilege, pEnabled)
                    If mZ.GetUserInfo(mMachineNumber, dwEnrollNumber, pName, pPassword, pPrivelege, pEnabled) Then
                        Dim c As New C_USERPARAM
                        c.AccessNo = dwEnrollNumber
                        mZ.GetStrCardNumber(c.CardNo)
                        c.Name = pName
                        c.Enabled = pEnabled
                        c.Priviledge = dwMachinePrivilege
                        c.Password = pPassword
                        cList.Add(c)
                    End If
                End While
            End If

            Return cList
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetUserInfoExt() As List(Of C_USERPARAMEXT)
        Try
            Dim cList As New List(Of C_USERPARAMEXT)

            Dim b As Boolean
            Dim pPrivelege As Integer
            Dim pName As String = ""
            Dim pPassword As String = ""
            Dim pEnabled As Boolean
            Dim dwEnrollNumber As Integer
            Dim dwMachinePrivilege As Integer
            Dim tmpData As String
            Dim name As String = ""

            b = mZ.ReadAllUserID(mMachineNumber)
            If b Then
                While mZ.GetAllUserInfo(mMachineNumber, dwEnrollNumber, pName, pPassword, dwMachinePrivilege, pEnabled)
                    If mZ.GetUserInfo(mMachineNumber, dwEnrollNumber, pName, pPassword, pPrivelege, pEnabled) Then
                        Dim c As New C_USERPARAMEXT
                        c.AccessNo = dwEnrollNumber
                        mZ.GetStrCardNumber(c.CardNo)
                        c.Name = pName
                        c.Enabled = pEnabled
                        c.Priviledge = dwMachinePrivilege
                        c.Password = pPassword

                        Dim length As Integer = 0

                        For i As Int32 = 0 To 9 Step 1
                            tmpData = ""
                            mZ.GetUserTmpExStr(mMachineNumber, CStr(dwEnrollNumber), i, 3, tmpData, length)
                            c.Fingers.Add(i, tmpData)
                        Next



                        cList.Add(c)
                    End If
                End While
            End If

            Return cList
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SSR_GetUserInfo(Optional action As Action(Of Int32) = Nothing) As List(Of C_USERPARAM)
        Try
            Dim cList As New List(Of C_USERPARAM)


            Dim pPrivelege As Integer
            Dim pName As String = ""
            Dim pPassword As String = ""
            Dim pEnabled As Boolean
            Dim dwEnrollNumber As String = ""
            Dim dwMachinePrivilege As Integer

            Dim name As String = ""

            Dim UserCount As Int32 = getRecordCount(eRecordCount.Registered_Users)
            Dim ii As Int32 = 0

            mZ.ReadAllUserID(mMachineNumber) '//read all the user information to the memory
            While mZ.SSR_GetAllUserInfo(mMachineNumber, dwEnrollNumber, pName, pPassword, dwMachinePrivilege, pEnabled)
                If mZ.SSR_GetUserInfo(mMachineNumber, dwEnrollNumber, pName, pPassword, pPrivelege, pEnabled) Then
                    Dim c As New C_USERPARAM
                    c.AccessNo = dwEnrollNumber
                    mZ.GetStrCardNumber(c.CardNo)
                    c.Name = pName
                    c.Enabled = pEnabled
                    c.Priviledge = dwMachinePrivilege
                    c.Password = pPassword
                    cList.Add(c)
                End If

                'RaiseEvent ShowStatus(Me, New C_EVENTHANDLERS.DeviceStatusEventArgs() With {.i = ii, .max = UserCount})
                If action IsNot Nothing Then action(ii)
                ii += 1
            End While
            '  End If

            Return cList
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SSR_GetUserInfoExt(clist As List(Of C_USERPARAM), mode As eUserExtMode, Optional action As Action(Of Int32) = Nothing) As List(Of C_USERPARAMEXT)
        Try
            Dim retlst As New List(Of C_USERPARAMEXT)
            Dim dwFaceIndex As Int32 = 50
            Dim tmpData As String = ""

            Dim hasFingerPrints As Int32 = getRecordCount(eRecordCount.FP)
            Dim hasFace As Int32 = getRecordCount(eRecordCount.Number_Face)

            Dim length As Integer = 0
            Dim ii As Int32 = 0

            mZ.ReadAllTemplate(mMachineNumber) '//read all the users' fingerprint templates to the memory
            For Each c As C_USERPARAM In clist
                Dim cc As New C_USERPARAMEXT
                With cc
                    .AccessNo = c.AccessNo
                    .CardNo = c.CardNo
                    .Enabled = c.Enabled
                    .Name = c.Name
                    .Password = c.Password
                    .Priviledge = c.Priviledge
                End With

                If mode = eUserExtMode.Fingers OrElse mode = eUserExtMode.Fingers_Face Then
                    If (hasFingerPrints > 0) Then
                        For i As Int32 = 0 To 9 Step 1
                            tmpData = ""
                            mZ.GetUserTmpExStr(mMachineNumber, CStr(c.AccessNo), i, 3, tmpData, length)
                            cc.Fingers.Add(i, tmpData)
                        Next
                    End If
                End If

                tmpData = ""
                length = 0

                If mode = eUserExtMode.Face OrElse mode = eUserExtMode.Fingers_Face Then
                    If hasFace > 0 Then
                        If Me.mZ.GetUserFaceStr(mMachineNumber, CStr(c.AccessNo), dwFaceIndex, tmpData, length) Then
                            cc.Faces.Add(dwFaceIndex, tmpData)
                        End If
                    End If
                End If

                'RaiseEvent ShowStatus(Me, New C_EVENTHANDLERS.DeviceStatusEventArgs() With {.i = ii, .max = clist.Count})
                If action IsNot Nothing Then action(ii)

                ii += 1

                retlst.Add(cc)
            Next

            Return retlst
        Catch ex As AccessViolationException
            Throw New Exception("The Device Could not perform this operation", ex)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SSR_DeleteUser(AccessNo As String) As Boolean
        Try
            mZ.SSR_DeleteEnrollData(mMachineNumber, AccessNo, 12)
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function DeleteUser(AccessNo As String, isSSR As Boolean) As Boolean
        Try

            If isSSR Then
                Return mZ.SSR_DeleteEnrollData(mMachineNumber, AccessNo, 12)
            Else
                'Return mZ.SSR_DeleteEnrollDataExt(mMachineNumber, AccessNo, 13)
                Return mZ.DeleteEnrollData(mMachineNumber, CInt(AccessNo), 1, 12)
            End If

            'Return mZ.DeleteEnrollData(mMachineNumber, AccessNo, 12, 0)
            'Return mZ.DeleteUserInfoEx(mMachineNumber, AccessNo)
            'Return mZ.SSR_DeleteEnrollData(mMachineNumber, AccessNo.ToString, 12)
            'mZ.SSR_DeleteEnrollDataExt(mMachineNumber, AccessNo, 13)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getNetOptions(ByVal StructNetOptionValue As eNetOptions) As String
        Return getSysOptions(StructNetOptionValue.ToString())
    End Function

    Public Function getSysOptions(ByVal StructNetOptionValue As String) As String
        Dim val As String = ""
        mZ.GetSysOption(mMachineNumber, StructNetOptionValue, val)
        Return val
    End Function

    Public Function getIpAddress() As String
        Dim Ipaddr As String = ""
        mZ.GetDeviceIP(1, Ipaddr)
        Return Ipaddr
    End Function

    Public Function UpdateDevice(file As String) As Boolean
        'Try
        '    Return mZ.UpdateFile(file)
        'Catch ex As Exception
        '    Throw ex
        'End Try
        Return Nothing
    End Function

    Public Function ShutdownDevice() As Boolean
        Try
            Return mZ.PowerOffDevice(mMachineNumber)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function UpdateFirmware(file As String) As Boolean
        Try
            Return mZ.UpdateFirmware(file)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function EnrollUser(e As C_USERPARAM) As Boolean
        Try
            Dim b As Boolean = False
            Dim accessNoStr As String = CStr(e.AccessNo)
            Dim accessNoInt As Integer = CInt(e.AccessNo)

            mZ.SetStrCardNumber(e.CardNo)

            If GetPlatform().Contains("TFT") Then
                b = mZ.SSR_SetUserInfo(mMachineNumber, accessNoStr, e.Name, e.Password, e.Priviledge, e.Enabled)
            Else
                b = mZ.SetUserInfo(mMachineNumber, accessNoInt, e.Name, e.Password, e.Priviledge, e.Enabled)
            End If

            'If b Then
            '    b = mZ.SetStrCardNumber(e.CardNo)
            'End If

            Return b
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function EnrollUserTmp(e As C_USERPARAMEXT, isSSR As Boolean) As Boolean
        Try
            Dim b As Boolean = False

            Dim accessNo As String = CStr(e.AccessNo)

            For Each ks As KeyValuePair(Of Int32, String) In e.Fingers

                If isSSR Then
                    b = mZ.SSR_SetUserTmpStr(mMachineNumber, accessNo, ks.Key, ks.Value)
                Else
                    b = mZ.SetUserTmpExStr(mMachineNumber, accessNo, ks.Key, 3, ks.Value)
                    If b = False Then
                        b = mZ.SSR_SetUserTmpStr(mMachineNumber, accessNo, ks.Key, ks.Value)
                    End If
                End If

                If b = False Then
                    Throw New Exception("Error Enrolling Finger Data : " & Me.GetError)
                End If
            Next

            Return b
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SetUserInfo(UserParam As C_USERPARAM, isSSR As Boolean) As Boolean
        Try
            Dim b As Boolean = False

            Dim accessNoStr As String = CStr(UserParam.AccessNo)
            Dim accessNoInt As Integer = CInt(UserParam.AccessNo)

            mZ.SetStrCardNumber(UserParam.CardNo)

            If isSSR Then
                b = mZ.SSR_SetUserInfo(mMachineNumber, accessNoStr, UserParam.Name.ToString(), UserParam.Password, CInt(UserParam.Priviledge), UserParam.Enabled)
            Else
                b = mZ.SetUserInfo(mMachineNumber, accessNoInt, UserParam.Name.ToString(), UserParam.Password, CInt(UserParam.Priviledge), UserParam.Enabled)
            End If

            Return b
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function SetUserInfo(ulist As List(Of C_USERPARAMEXT)) As List(Of Int32)
        Try
            Dim EnrollCount As Integer = 0
            Dim UnEnrollCount As Integer = 0
            Dim isSSR As Boolean = Me.GetPlatform().Contains("TFT")

            For Each f As C_USERPARAMEXT In ulist

                Me.DeleteUser(CStr(f.AccessNo), isSSR) ' Remove Trash data`
                If Me.EnrollUser(f) Then 'Enroll User generic data
                    Me.EnrollUserTmp(f, isSSR)
                    EnrollCount = 1
                Else
                    UnEnrollCount = 1
                End If
            Next

            Dim b As New List(Of Int32)
            b.AddRange(New Int32() {EnrollCount, UnEnrollCount, ulist.Count})
            Return b
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function DeleteUserInfo(ulist As List(Of C_USERPARAM)) As List(Of Int32)
        Try
            Dim DeleteCount As Integer = 0
            Dim UnDeleteCount As Integer = 0
            Dim isSSR As Boolean = Me.GetPlatform().Contains("TFT")

            For Each f As C_USERPARAM In ulist

                If Me.DeleteUser(CStr(f.AccessNo), isSSR) Then 'Enroll User generic data
                    DeleteCount = 1
                Else
                    UnDeleteCount = 1
                End If
            Next

            Dim b As New List(Of Int32)
            b.AddRange(New Int32() {DeleteCount, UnDeleteCount, ulist.Count})
            Return b
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetError() As String
        Dim i As Int32
        Me.mZ.GetLastError(i)
        Return GetErrorCode(CType(i, eErrorCode))
    End Function

    Public Function BeginBatchUpdate() As Boolean
        Return mZ.BeginBatchUpdate(mMachineNumber, 1)
    End Function

    Public Sub SetDisplay(ByVal msg As String)
        mZ.WriteLCD(0, 0, msg)
    End Sub

    Public Sub ClearDisplay()
        mZ.ClearLCD()
    End Sub

    Public Function GetErrorCode(i As eErrorCode) As String
        Return i.ToString.Replace("_", " ")
    End Function

    'Public Function EnableKeyboard(enble As Boolean) As Boolean

    'End Function

#End Region

#Region " OVERRIDES "

    Protected Overrides Sub Finalize()
        Try
            Disconnect()
            mZ = Nothing
        Catch ex As Exception
        Finally
            MyBase.Finalize()
        End Try
    End Sub

#End Region

#Region " ENUMERATIONS "

    Public Enum eRecordCount As Int32
        Administrator = 1
        Registered_Users = 2
        FP = 3
        Passwords = 4
        Operation_records = 5
        Attendance_records = 6
        Finger_Print_Capacity = 7
        User_Capacity = 8
        Attendance_Record_Capacity = 9
        Residual_FingerPrint_Capacity = 10
        Residual_User_Capacity = 11
        Residual_Attendance_Record_Capacity = 12
        Number_Face = 21
        Face_Capacity = 22
    End Enum

    Public Enum eConnectType As Int32
        Net
        Com
        RS232
    End Enum

    Public Enum eNetOptions As Int32
        GATEIPAddress
        NetMask
    End Enum

    Public Enum eClearFlag As Int32
        Attendance_Record = 1
        Fingerprint_Template = 2
        None = 3
        Operation_Record = 4
        User_Information = 5
    End Enum

    Public Enum eVerifyStyle As Int32
        FP_OR_PW_OR_RF = 0
        FP = 1
        PIN = 2
        PW = 3
        RF = 4
        FP_OR_PW = 5
        FP_OR_RF = 6
        PW_OR_RF = 7
        PIN_AND_FP = 8
        FP_AND_PW = 9
        FP_AND_RF = 10
        PW_AND_RF = 11
        FP_AND_PW_AND_RF = 12
        PIN_AND_FP_AND_PW = 13
        FP_AND_RF_OR_PIN = 14
    End Enum

    Public Enum eVerifyStyle_TFT As Int32
        FPorPWorRF = 128
        FP = 129
        PIN = 130
        PW = 131
        RF = 132
        FPandRF = 133
        FPorPW = 134
        FPorRF = 135
        PWorRF = 136
        PINandFP = 137
        FPandPW = 138
        PWandRF = 139
        FPandPWandRF = 140
        PINandFPandPW = 141
        FPandRForPIN = 142
    End Enum

    Public Enum eUserExtMode As Int32
        Fingers = 0
        Face = 1
        Fingers_Face = 2
    End Enum

    Public Enum eErrorCode As Int32
        Operation_failed_or_data_not_exist = -100
        Transmitted_data_length_is_incorrect = -10
        Data_already_exists = -5
        Space_is_not_enough = -4
        Error_size = -3
        Error_in_file_read_write = -2
        SDK_is_not_initialized_and_needs_to_be_reconnected = -1
        Data_not_found_or_data_repeated = 0
        Operation_is_correct = 1
        Parameter_is_incorrect = 4
        Error_in_allocating_buffer = 101
    End Enum

#End Region

#Region " SysOptions "

    Public Const SYSOPT_ZKPVersion As String = "~ZKFPVersion"
    Public Const Must1To1 As String = "Must1To1"

#End Region

#Region " PROPERTIES "

    Public ReadOnly Property pID As Int32
        Get
            Return mID
        End Get
    End Property

    Public Property pIP() As String
        Get
            Return mIP
        End Get
        Set(ByVal Value As String)
            mIP = Value
        End Set
    End Property

    Public Property pPort() As Integer
        Get
            Return mPort
        End Get
        Set(ByVal Value As Integer)
            mPort = Value
        End Set
    End Property

    Public ReadOnly Property pMachineNumber() As Integer
        Get
            Return mMachineNumber
        End Get
    End Property

    Public Property pDeviceName As String
        Get
            Return mDeviceName
        End Get
        Set(value As String)
            Me.mDeviceName = value
        End Set
    End Property

    Public ReadOnly Property pIsConnected As Boolean
        Get
            Return Me.mConnected
        End Get
    End Property

    '- Com Port

    Public WriteOnly Property pComPort As Integer
        Set(value As Integer)
            Me.mComPort = value
        End Set
    End Property

    Public WriteOnly Property pBaudRate As Integer
        Set(value As Integer)
            Me.mBaudRate = value
        End Set
    End Property

    '- Logs

    ''' <summary>
    ''' LJ 20130912
    ''' This will be Initiated when void Time log 2 is executed
    ''' </summary>
    ''' <remarks></remarks>
    Private mFullLogs() As String
    Public ReadOnly Property pFullLogs() As String()
        Get
            Return Me.mFullLogs
        End Get
    End Property

    ''' <summary>
    ''' LJ 20130912
    ''' This will be Initiated when void Time log 2 is executed
    ''' </summary>
    ''' <remarks></remarks>
    Private mPartialLogs() As String
    Public ReadOnly Property pPartialLogs As String()
        Get
            Return Me.mPartialLogs
        End Get
    End Property

    Public Property pMacAddress As String
        Get
            Dim s As String = ""
            Me.mZ.GetDeviceMAC(mMachineNumber, s)
            Return s
        End Get
        Set(value As String)
            Me.mZ.SetDeviceMAC(mMachineNumber, value)
        End Set
    End Property

    Public Property pNetOptions(opt As eNetOptions) As String
        Get
            Return getNetOptions(opt)
        End Get
        Set(value As String)
            setNetOptions(opt, value)
        End Set
    End Property

    Public Function GetUserFingerEx(AccessNo As String, index As Int32) As String
        Dim tmpdata As String = ""
        Dim length As Int32 = 0
        mZ.GetUserTmpExStr(mMachineNumber, AccessNo, index, 3, tmpdata, length)
        Return tmpdata
    End Function

    Public Function GetUserFinger(AccessNo As Integer, index As Int32) As String
        Dim tmpdata As String = ""
        Dim length As Int32 = 0
        mZ.GetUserTmpStr(mMachineNumber, CInt(AccessNo), index, tmpdata, length)
        Return tmpdata
    End Function

    Public Property Tag As Object

#End Region

#Region " INNER CLASSES "

    Class Connectify
        Implements IDisposable

        Public Delegate Sub OnConnect() '(sender As C_DEVICE, e As EventArgs)
        Public Delegate Sub OnDisconnected() '(sender As C_DEVICE, e As EventArgs)

        Public Event Connect As OnConnect
        Public Event DisConnect As OnDisconnected

        Public Sub StartConnect(dev As C_DEVICE, Optional ByVal ConnectType As eConnectType = eConnectType.Net)
            Try
                'Try to connect to device
                SyncLock dev
                    Select Case ConnectType
                        Case eConnectType.Net
                            If CanPing(dev.mIP) Then
                                dev.mConnected = dev.mZ.Connect_Net(dev.mIP, dev.mPort)
                            Else
                                dev.mConnected = False '(dev, EventArgs.Empty)
                            End If
                        Case eConnectType.Com : dev.mConnected = dev.mZ.Connect_Com(dev.mComPort, dev.mMachineNumber, dev.mBaudRate)
                        Case eConnectType.RS232 : dev.mConnected = dev.mZ.Connect_USB(dev.mMachineNumber)
                    End Select

                    If dev.mConnected = True Then
                        dev.RegEvents(dev.mMachineNumber, 65535)
                        RaiseEvent Connect() '(dev, EventArgs.Empty)
                    Else
                        RaiseEvent DisConnect() '(dev, EventArgs.Empty)
                    End If
                End SyncLock
            Catch

            End Try
        End Sub

        Function CanPing(ip As String) As Boolean
            Dim tryCount = 0

            Using png As New Net.NetworkInformation.Ping
                Dim pngReply As Net.NetworkInformation.PingReply = png.Send(ip, 10000)

                While pngReply.Status <> Net.NetworkInformation.IPStatus.Success

                    If tryCount >= 1 Then
                        'Throw New Exception("Cannot Ping Device")
                        Return False
                    End If

                    pngReply = png.Send(ip, 10000)

                    Threading.Thread.Sleep(1000)
                    tryCount += 1
                End While
            End Using

            Return True
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose

            GC.Collect()
            GC.SuppressFinalize(Me)

            'Try
            '    Threading.Thread.CurrentThread.Abort()
            'Catch
            'End Try

        End Sub

    End Class

#End Region

End Class
