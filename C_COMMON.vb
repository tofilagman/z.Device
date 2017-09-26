Public NotInheritable Class C_COMMON

    Public Shared Sub RegisterZkemkeeper(Path As String)

        'Dim fileName As String = "regZkem.bat"
        Dim zkemFile As String = "zkemkeeper.dll"
        Dim filepath As String = Path 'Environment.CurrentDirectory
        Dim arraytext As String = ""
        Try
            'arraytext &= "cd " & filepath & vbCrLf
            'arraytext &= Strings.Left(filepath, 1) & ":" & vbCrLf
            arraytext &= "regsvr32 """ & IO.Path.Combine(Path, zkemFile) & """ /s"

            '    If System.IO.File.Exists(filepath & "\" & zkemFile) Then
            '        System.IO.File.Delete(filepath & "\" & fileName)
            '        System.IO.File.AppendAllText(filepath & "\" & fileName, arraytext)
            '    End If

            '    If System.IO.File.Exists(filepath & "\" & fileName) Then

            '        Dim startInfo As New ProcessStartInfo(fileName)
            '        startInfo.WindowStyle = ProcessWindowStyle.Hidden
            '        Process.Start(startInfo)
            '    End If
            If IO.File.Exists(IO.Path.Combine(Path, zkemFile)) = True Then
                Dim pri As New ProcessStartInfo("cmd", "/c " + arraytext)
                Dim pr As New Process
                pr.StartInfo = pri
                pr.Start()
                pr.WaitForExit()
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Shared Function GetDataSetFromUserParameters(usrlst As List(Of C_USERPARAMEXT)) As DataSet
        Dim ds As New DataSet("UserInfo")

        Dim dt As New DataTable("User")
        dt.Columns.Add("Name", GetType(String))
        dt.Columns.Add("AccessNo", GetType(String))
        dt.Columns.Add("CardNo", GetType(String))
        dt.Columns.Add("Password", GetType(String))
        dt.Columns.Add("Priviledge", GetType(Int32))
        dt.Columns.Add("Enabled", GetType(Boolean))

        Dim dt2 As New DataTable("Fingers")
        dt2.Columns.Add("AccessNo", GetType(String))
        dt2.Columns.Add("FingerID", GetType(Int32))
        dt2.Columns.Add("Template", GetType(Object))

        For Each s As C_USERPARAMEXT In usrlst
            Dim dr As DataRow = dt.NewRow
            dr("Name") = s.Name
            dr("AccessNo") = s.AccessNo
            dr("CardNo") = s.CardNo
            dr("Password") = s.Password
            dr("Priviledge") = s.Priviledge
            dr("Enabled") = s.Enabled
            dt.Rows.Add(dr)
        Next

        For Each s As C_USERPARAMEXT In usrlst
            For Each f In s.Fingers
                For Each fngr As KeyValuePair(Of Int32, String) In s.Fingers
                    If Not IsNothing(fngr.Value) Then
                        Dim idr As DataRow = dt2.NewRow
                        idr("AccessNo") = s.AccessNo
                        idr("FingerID") = fngr.Key
                        idr("Template") = IIf(IsNothing(fngr.Value), DBNull.Value, fngr.Value)
                        dt2.Rows.Add(dt2)
                    End If
                Next
            Next
        Next

        ds.Tables.Add(dt)
        ds.Tables.Add(dt2)

        Return ds
    End Function

    Public Shared Function GetSDKVersion() As String
        Return System.Reflection.Assembly.GetAssembly(GetType(zkemkeeper.IZKEM)).GetName().Version.ToString
    End Function

End Class
