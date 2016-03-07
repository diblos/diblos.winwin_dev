Public Class MidnightClass 'AVL 4G TRACKER MIGNIGHT 5 MINUTES FILTER

    Private _trackerId As String

    Public Function doProceed(ByVal data As String) As Boolean
        Dim ti As Date = Now
        If ti.Hour >= 0 And ti.Hour <= 4 Then 'CHECK TIME WINDOW
            If GetACCStatus(data) = ACCStatus.ON Then 'CHECK ACC STATUS
                Return True
            Else
                _trackerId = GetTrackerID(data)
                If CheckGPSDatetime(_trackerId, data) Then 'TRUE (5 MINUTES CHECKING)
                    Return True
                Else
                    Return False
                End If
            End If
        Else
            Return False
        End If
    End Function

#Region "PRIVATE METHODS"

    Private Function GetACCStatus(ByVal _data As String) As ACCStatus
        Try
            If ConvertHexToBinary(_data.Substring(12, 2)).ToString.Substring(1, 1) = "1" Then
                'ACC ON
                Return ACCStatus.ON
            Else
                'ACC OFF
                Return ACCStatus.OFF
            End If

        Catch ex As Exception
            Return ACCStatus.OFF
        End Try
    End Function

    Private Function GetTrackerID(ByVal _data As String) As String
        Try
            Dim _queue As String = _data.Substring(91, 1)
            'last 5 digit
            Dim tracker As String = _data.Substring(79, 1) & _data.Substring(82, 1) & _data.Substring(85, 1) & _
            _data.Substring(88, 1) & _queue

            Return "NR09G" & tracker
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Private Enum ACCStatus
        [ON]
        [OFF]
    End Enum

#Region "BAKHTIAR TEST: 300 SECONDS IGNORE DATA"
    Public Const DATETIME_FORMAT = "yyyy-MM-dd HH:mm:ss"
    Public dec0F As Integer = Convert.ToInt32("0x0F", 16)
    Public dec1F As Integer = Convert.ToInt32("0x1F", 16)
    Public dec3F As Integer = Convert.ToInt32("0x3F", 16)

    Private Enum IncomingStatus
        Inserted
        Updated
        Ignored
    End Enum

    Private Function CheckGPSDatetime(ByVal TrackerID As String, ByVal RawData As String) As Boolean

        Try
            Select Case updateDT(_trackerId, GPSDatetime(RawData))
                Case IncomingStatus.Ignored
                    Return False
                Case IncomingStatus.Inserted
                    Return True
                Case IncomingStatus.Updated
                    Return True
                Case Else
                    Return True
            End Select
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function updateDT(ByVal ID As String, ByVal Timestamp As Date) As IncomingStatus

        Dim AVL As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("AVL")

        Dim Sql As New System.Text.StringBuilder

        Sql.Append(" SELECT MAX([GPSDatetime]) AS GPSDatetime ")
        Sql.Append(" FROM [avltmp] where [Tracker_ID] = '" & ID & "'")

        Dim Command As System.Data.Common.DbCommand = AVL.GetSqlStringCommand(Sql.ToString)
        Dim dTime As String = String.Empty
        Try
            dTime = AVL.ExecuteScalar(Command)
            If DateDiff(DateInterval.Second, CDate(dTime), Timestamp) < 300 Then '300 SECONDS = 5 MINUTES
                Return IncomingStatus.Ignored
            Else
                'UPDATE HERE
                Sql.Remove(0, Sql.Length)
                Sql.Append(" UPDATE [avltmp] SET [GPSDatetime] ='" & Timestamp.ToString(DATETIME_FORMAT) & "' ")
                Sql.Append(" WHERE [Tracker_ID] = '" & ID & "' ")
                Command = AVL.GetSqlStringCommand(Sql.ToString)
                AVL.ExecuteNonQuery(Command)
                Return IncomingStatus.Updated
            End If
        Catch ex As Exception
            'INSERT HERE
            Sql.Remove(0, Sql.Length)

            Sql.Append("INSERT INTO [avltmp] ([GPSDatetime],Tracker_ID) VALUES (")
            Sql.Append("'" & Timestamp.ToString(DATETIME_FORMAT) & "', '" & ID & "')")
            Command = AVL.GetSqlStringCommand(Sql.ToString)
            AVL.ExecuteNonQuery(Command)
            Return IncomingStatus.Inserted
        Finally
            Sql = Nothing
            Command.Connection.Close()
            Command.Dispose()
            GC.Collect()
        End Try

    End Function

    Private Function GPSDatetime(ByVal _data As String) As Object
        Dim splitData As String()

        Try
            splitData = _data.Split(":")

            If (splitData Is Nothing) Then
                Return Nothing
            End If

            If (splitData.Length <= 0) Then
                Return Nothing
            End If

            Dim _strGPSDatetime As String = splitData(20) + splitData(19) + splitData(18) + splitData(17)

            Return ConvertGPSDatetime(_strGPSDatetime)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function ConvertGPSDatetime(ByVal _strGPSDatetime As String) As Date
        Dim gpsNewDatetime As String
        Dim dec As Integer
        Dim year As Integer
        Dim month As Integer
        Dim day As Integer
        Dim hour As Integer
        Dim minute As Integer
        Dim second As Integer
        Dim dtNow = Now()

        Try
            dec = ConvertHexToDec(_strGPSDatetime)

            year = dec >> 26
            month = (dec >> 22) And dec0F
            day = (dec >> 17) And dec1F
            hour = (dec >> 12) And dec1F
            minute = (dec >> 6) And dec3F
            second = dec And dec3F


            year = Convert.ToInt64(Date.Now.Year.ToString.Substring(0, 2) & year)

            'LIVE Uncomment the line below
            gpsNewDatetime = Format(New DateTime(year, month, day, hour, minute, second, 0), DATETIME_FORMAT)
            'LOCAL Uncomment the line below
            'gpsNewDatetime = Format(New DateTime(dtNow.Year, dtNow.Month, dtNow.Day, hour, minute, second, 0), DATETIME_FORMAT)

            Dim _gpsDatetime As Date = DateTime.SpecifyKind(Convert.ToDateTime(gpsNewDatetime).ToLocalTime, DateTimeKind.Local)

            'If (Date.Compare(_gpsDatetime.Date, Now().Date) > 0) Then
            '    Return Nothing
            'End If

            Return _gpsDatetime
        Catch fex As FormatException
            Return Nothing
        Catch aex As ArgumentOutOfRangeException
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function ConvertHexToDec(ByVal hexData As String) As Integer
        Try
            Return Convert.ToInt32(hexData, 16)
        Catch ex As Exception
            Return 0
        End Try
    End Function
#End Region

    Private Function ConvertHexToBinary(ByVal hexData As String) As Integer
        Dim decData As Integer

        Try
            decData = ConvertHexToDec(hexData)
            Return Convert.ToString(decData, 2)
        Catch ex As Exception
            Return 0
        End Try
    End Function

#End Region

End Class
