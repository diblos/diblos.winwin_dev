Imports Itramas.Itac.Helper
Imports System.Net.NetworkInformation

Public Class Form1

    Dim nSize As New Size(600, 400)
    Dim ListBoxMaxCount As Integer = 9999
    Dim CLEARLIST As Boolean = True

    '=====================================================================================
    'SPECIAL FUNCTIONALITY TEST
    Dim smsQueue As QueueReader
    Dim _timer As System.Threading.Timer
    Dim _timerCallback As System.Threading.TimerCallback = AddressOf TimerIsTicking
    '=====================================================================================
    Dim counter As Integer = 0

    Public Delegate Sub AddItemsToListBoxDelegate( _
                     ByVal ToListBox As ListBox, _
                     ByVal AddText As String)

    Public Delegate Sub CaptionFormDelegate( _
        ByVal mForm As Form1, _
        ByVal txt As String)

    Private Sub AddItemsToListBox(ByVal ToListBox As ListBox, _
                                 ByVal AddText As String)
        If ToListBox.Items.Count > ListBoxMaxCount And CLEARLIST = True Then
            ToListBox.Items.Clear()
        End If

        ToListBox.Items.Add(AddText)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, True)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, False)
    End Sub

    Public Sub lstMsgs(ByVal item As Object, Optional ByVal source As Object = Nothing)
        If (ListBox1.InvokeRequired) Then
            ListBox1.Invoke( _
                    New AddItemsToListBoxDelegate(AddressOf AddItemsToListBox), _
                    New Object() {ListBox1, CStr(item)})
        Else
            If Me.ListBox1.Items.Count > ListBoxMaxCount And CLEARLIST = True Then
                ListBox1.Items.Clear()
            End If

            Me.ListBox1.Items.Add(CStr(item))
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, True)
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, False)
        End If

    End Sub

    Private Sub recapFormCaption(ByVal mForm As Form1, ByVal txt As String)
        Me.Text = txt
    End Sub

    Private Sub ReCaptionForm(ByVal txt As String)
        If (Me.InvokeRequired) Then
            Me.Invoke( _
                    New CaptionFormDelegate(AddressOf recapFormCaption), _
                    New Object() {Me, txt})
        Else
            Me.Text = txt
        End If
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If IsNothing(smsQueue) = False Then smsQueue.Dispose()
            If IsNothing(_timer) = False Then _timer.Dispose()
        Catch ex As Exception
            'lstMsgs(ex.Message)
            MsgBox(ex.Message)
        Finally
            End
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Initialization()
        Me.Text = APPLICATION_NAME
        Me.Size = nSize
        Me.MaximumSize = nSize

        With Me.ListBox1
            .BackColor = Color.Black
            .ForeColor = Color.Lime
        End With


        'lstMsgs(Now.ToString("yyyy-MM-dd 03:00:00"))
        'Test1()
        'Test2()
        'Test3()
        'Test4()
        'Test5()
        'Test6()
        'lstMsgs(GetFirstInterval)
        'initQueue()
        'initThreadTimer()



        'CheckAVLDatetime()
        'CheckAVLFrequencies()

        'TestScreen()

        'lstMsgs(CheckLastDatetime("NR09G20844x"))

        'TestDateDiff()
        'TestPing()

        'TestDatatable2Object()

        'TestTable()

        '===================================================================
        'RebuildTableIndexRPT("AVL_DataLoss")
        'Dim k As New System.Threading.Thread(AddressOf RebuildTableIndexRPT)
        'k.Start("AVL_DataLoss")
        'k.Start("AVL_Data_Hist_03")
        '===================================================================
        'Dim kk As New AVLEngineClass
        'lstMsgs(kk.GetLength)
        'lstMsgs(kk.GetString)

        'ParseData(kk.GetString)
        '===================================================================
        'Dim engine As New MidnightClass
        'Dim status As Boolean = engine.doProceed("22:00:08:00:F4:40:00:00:00:1A:36:CB:42:BC:93:4A:40:D5:79:30:39:4E:52:30:39:47:33:30:30:30:31:00:00:00:22:00:08:00:84:40:00:00:00:30:36:CB:42:75:93:4A:40:EA:79:30:39:4E:52:30:39:47:33:30:30:30:31:00:00:00:22:00:08:00:84:40:00:00:00:30:36:CB:42:75:93:4A:40:02:7A:30:39:4E:52:30:39:47:33:30:30:30:31:00:00:00:22:00:08:00:84:40:00:00:00:30:36:CB:42:A3:92:4A:40:16:7A:30:39:4E:52:30:39:47:33:30:30:30:31:00:00:00:22:00:08:00:84:40:00:00:00:46:36:CB:42:D1:91:4A:40:2A:7A:30:39:4E:52:30:39:47:33:30:30:30:31:00:00:00:22:00:08:00:84:40:00:00:00:46:36:CB:42:8A:91:4A:40:42:7A:30:39:4E:52:30:39:47:33:30:30:30:31:00:00:00")

        'lstMsgs(status)
        '===================================================================
        'Dim xstr As String = "ASAL"
        'Dim t As New Test
        't.Execute(xstr)
        'lstMsgs(xstr)
        '===================================================================
        'POLY
        'DayAtTheAmusementPark()

        'INHERITS
        Dim oLine As LineDelim = New LineDelim()

        oLine.Line = "aku budak minang"
        lstMsgs(oLine.GetWord())
        lstMsgs(oLine.aGetWord())
    End Sub

#Region "PrivateTest"

    Private Sub DayAtTheAmusementPark()
        Dim oRollerCoaster As New TheRollerCoaster
        Dim oMerryGoRound As New TheMerryGoRound
        Call GoOnRide(oRollerCoaster)
        Call GoOnRide(oMerryGoRound)
    End Sub
    Private Sub GoOnRide(ByVal oRide As Object)
        oRide.Ride()
    End Sub

    Private Sub Test1()
        Dim nValue As Integer = 1
        If nValue <> vbNull Or nValue <> 0 Then
            lstMsgs("YES")
        Else
            lstMsgs("NO")
        End If
    End Sub

    Private Sub Test2()
        Dim _data As String = "RESETSEGMENT"
        If InStr(_data, "RESETSEGMENT") = 0 Then
            lstMsgs("NO" & vbTab & InStr(_data, "RESETSEGMENT"))
        Else
            lstMsgs("YES" & vbTab & InStr(_data, "RESETSEGMENT"))
        End If
    End Sub

    Private Sub Test3() 'Check Negative Value
        Dim DATA As Integer = 4
        Dim BUS_LATE_THRESHOLD As Integer = -5

        lstMsgs(DATA < BUS_LATE_THRESHOLD)
        lstMsgs(DATA > BUS_LATE_THRESHOLD)

        lstMsgs(BUS_LATE_THRESHOLD * -1)

    End Sub

    Private Sub Test4() 'Run After Midnight
        Dim batchStartTime As String
        Dim dtBatchStartTime As DateTime

        batchStartTime = Configuration.ConfigurationManager.AppSettings("BatchStartTime")

        lstMsgs(batchStartTime)

        dtBatchStartTime = DateTime.Parse(Now.ToString("yyyy-MM-dd " & batchStartTime))
        If DateTime.Parse(Now.ToString("yyyy-MM-dd " & batchStartTime)) < Now Then
            dtBatchStartTime = CDate(Now.AddDays(1).ToString("yyyy-MM-dd " & batchStartTime))
        End If

        lstMsgs(dtBatchStartTime.ToString(DEFAULT_DATESTRING))

        Dim firstInterval = (DateDiff(DateInterval.Second, Now, dtBatchStartTime)) * 1000

        lstMsgs(firstInterval)
    End Sub

    Private Sub Test5()
        Dim batchStartTime = Configuration.ConfigurationManager.AppSettings("BatchStartTime")
        'lstMsgs(New TimerClass(batchStartTime).GetInterval)
        lstMsgs(New TimerClass(batchStartTime).GetDate)
    End Sub

    Private Sub Test6()
        Dim firstInterval = (DateDiff(DateInterval.Second, Now, CDate(Now.AddHours(1).ToString("yyyy-MM-dd HH:00")))) * 1000
        lstMsgs(firstInterval)
        lstMsgs(DateAdd(DateInterval.Second, firstInterval / 1000, Now).ToString(DEFAULT_DATESTRING))

        If firstInterval > 1800000 Then '30 minutes
            'dtGetDate = Now()
            'LOGGER.WriteToLog(Me.ServiceName, ": Process Purging Start")
            'Purge_Data()
            'LOGGER.WriteToLog(Me.ServiceName, ": Process Purging Complete")
        End If

    End Sub

    Private Sub initQueue()
        smsQueue = New QueueReader(COMMAND_QUEUE, AddressOf ProcessQueueCommand)
        smsQueue.StartReceive()
    End Sub

    Private Sub initThreadTimer()
        '_timer = New System.Threading.Timer(_timerCallback, Nothing, 0, 30000)
        _timer = New System.Threading.Timer(_timerCallback, Nothing, 0, 1000)
    End Sub

    Private Sub ProcessQueueCommand(ByVal qData As String)
        Try
            lstMsgs("Q Received " & qData.ToString)
        Catch ex As Exception
            lstMsgs(ex.Message)
        End Try
    End Sub

    Private Sub TimerIsTicking()
        lstMsgs(Now.ToString(DEFAULT_DATESTRING))
        counter += 1
        If counter = 10 Then
            PauseTimer(_timer)

            System.Threading.Thread.Sleep(20000)

            ResumeTimer(_timer)
        End If
    End Sub

    Private Function CheckLastDatetime(ByVal TrackerID As String) As Boolean

        Me.Text = TrackerID
        'CLEARLIST = False

        Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("AVL")

        Dim Sql As New System.Text.StringBuilder

        Sql.Append(" SELECT  MAX([Received_Datetime])[Received_Datetime]")
        Sql.Append(" FROM [avldb].[dbo].[avl_raw] with (nolock) where [Tracker_ID] = '" & TrackerID & "' ")

        Dim Command As System.Data.Common.DbCommand = db.GetSqlStringCommand(Sql.ToString)

        Try
            Dim s = db.ExecuteScalar(Command)
            lstMsgs(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

        Command.Connection.Close()
        Command.Dispose()
        GC.Collect()

    End Function

    Private Sub TestDateDiff()
        lstMsgs(DateDiff(DateInterval.Second, CDate("2015-11-27 10:19"), Now))
        lstMsgs(DateDiff(DateInterval.Second, CDate("2015-11-27 10:19"), CDate("2015-11-27 10:00")))
        lstMsgs(DateDiff(DateInterval.Second, CDate("2015-11-27 10:19:00"), CDate("2015-11-27 10:19:20")))
    End Sub

    Private Sub TestTable()
        Dim TrackerID As String = "NR09G20844"
        Me.Text = TrackerID

        Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("TAVL")

        Dim Sql As New System.Text.StringBuilder

        Sql.Append(" SELECT MAX([GPSDatetime]) AS GPSDatetime ")
        Sql.Append(" FROM [avltmp] where [Tracker_ID] = '" & TrackerID & "'")

        Dim Command As System.Data.Common.DbCommand = db.GetSqlStringCommand(Sql.ToString)
        Dim dTime As String = String.Empty
        Try
            dTime = db.ExecuteScalar(Command)
            If DateDiff(DateInterval.Second, CDate(dTime), Now) < 19 Then
                lstMsgs("IGNORED")
            Else
                'UPDATE HERE
                Sql.Remove(0, Sql.Length)
                Sql.Append(" UPDATE [avltmp] SET [GPSDatetime] ='" & Now.ToString(DATETIME_FORMAT) & "' ")
                Sql.Append(" WHERE [Tracker_ID] = '" & TrackerID & "' ")
                Command = db.GetSqlStringCommand(Sql.ToString)
                db.ExecuteNonQuery(Command)
                lstMsgs("UPDATED")
            End If
        Catch ex As Exception
            'INSERT HERE
            Sql.Remove(0, Sql.Length)

            Sql.Append("INSERT INTO [avltmp] ([GPSDatetime],Tracker_ID) VALUES (")
            Sql.Append("'" & Now.ToString(DATETIME_FORMAT) & "', '" & TrackerID & "')")
            Command = db.GetSqlStringCommand(Sql.ToString)
            db.ExecuteNonQuery(Command)
            lstMsgs("INSERTED")
        End Try

        Command.Connection.Close()
        Command.Dispose()
        GC.Collect()
    End Sub

    Private Sub TestDatatable2Object()
        Dim TrackerID As String = "NR09G20844"
        Me.Text = TrackerID
        'CLEARLIST = False

        Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("AVL")

        Dim Sql As New System.Text.StringBuilder

        Sql.Append(" SELECT Received_Datetime,Tracker_ID ")
        Sql.Append(" FROM [avldb].[dbo].[avl_raw] where [Tracker_ID] = '" & TrackerID & "' ")
        Sql.Append(" And DatePart(Hour, [Received_Datetime]) = 19")
        'Sql.Append(" and DATEPART(MINUTE, [ReceivedDatetime])=0")
        Sql.Append(" order by [Received_Datetime]")

        Dim Command As System.Data.Common.DbCommand = db.GetSqlStringCommand(Sql.ToString)
        Dim dt As DataTable = db.ExecuteDataSet(Command).Tables(0)
        Command.Connection.Close()

        For Each nRow As DataRow In dt.Rows
            lstMsgs(nRow("Received_Datetime") & vbTab & nRow("Tracker_ID"))
        Next

        dt.Dispose()
        Command.Dispose()

        GC.Collect()
    End Sub

    Private Function GetListByDataTable(ByVal dt As DataTable) As List(Of Object)
        'Dim reult = (From rw In dt.AsEnumerable(), New With { _
        '	Key .Name = Convert.ToString(rw("Name")), _
        '	Key .ID = Convert.ToInt32(rw("ID")) _
        '}).ToList()

        '       Return reult.ConvertAll(Of Object)(Function(o) DirectCast(o, Object))
        Return Nothing
    End Function

#Region "PING TEST"
    Private Sub TestPing()
        Dim ip() As String = {"192.168.8.42", "192.168.8.33", "192.168.8.23", "google.com", "8.8.8.8", "llm.gov.my", "kwsp.gov.my"}
        For Each address In ip
            Dim k As New System.Threading.Thread(AddressOf PingMethod)
            k.Start(address)
        Next
    End Sub

    Private Sub PingMethod(ByVal address As String)
        Dim p As New Ping
        Try
            lstMsgs(address & vbTab & p.Send(address).Status.ToString)
        Catch ex As Exception
            lstMsgs(address & vbTab & ex.InnerException.Message)
        End Try
        p.Dispose()
        GC.Collect()
    End Sub
#End Region

#Region "Rebuild Index"
    Private Sub RebuildTableIndexRPT(ByVal strTable As String)
        Dim startTime As Date = Now
        lstMsgs("RebuildTableIndexRPT: " & strTable)
        Dim RPT As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
       Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("REPORT")

        Dim sql, strAlter As String
        Dim intI, intRowcount, elapsedSeconds As Integer
        Dim dtExecute, dtComplete As DateTime
        Dim ds As DataSet
        Dim command As System.Data.Common.DbCommand

        Try
            sql = "SELECT ('ALTER INDEX [' + ix.name + '] ON [' + s.name + '].[' + t.name + '] ' + " & _
                           "CASE WHEN ps.avg_fragmentation_in_percent > 30 THEN 'REBUILD' ELSE 'REORGANIZE' END + " & _
                           "CASE WHEN pc.partition_count > 1 THEN ' PARTITION = ' + cast(ps.partition_number as nvarchar(max)) ELSE '' END) AS AlterString " & _
                    "FROM   sys.indexes AS ix " & _
                           "INNER JOIN sys.tables t " & _
                               "ON t.object_id = ix.object_id " & _
                           "INNER JOIN sys.schemas s " & _
                               "ON t.schema_id = s.schema_id " & _
                           "INNER JOIN (SELECT object_id, index_id, avg_fragmentation_in_percent, partition_number " & _
                                       "FROM sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL, NULL, NULL)) ps " & _
                               "ON t.object_id = ps.object_id AND ix.index_id = ps.index_id " & _
                           "INNER JOIN (SELECT object_id, index_id, COUNT(DISTINCT partition_number) AS partition_count " & _
                                        "FROM sys.partitions " & _
                                        "GROUP BY object_id, index_id) pc " & _
                               "ON t.object_id = pc.object_id AND ix.index_id = pc.index_id " & _
                    "WHERE  ps.avg_fragmentation_in_percent > 5 " & _
                           "AND ix.name IS NOT NULL " & _
                           "AND t.name = @TABLE " & _
                    "ORDER BY ps.avg_fragmentation_in_percent"

            command = RPT.GetSqlStringCommand(sql)
            RPT.AddInParameter(command, "@TABLE", DbType.String, strTable)
            command.CommandTimeout = 0
            ds = RPT.ExecuteDataSet(command)
            command.Connection.Close()

            intRowcount = ds.Tables(0).Rows.Count

            If intRowcount > 0 Then
                For intI = 0 To intRowcount - 1
                    strAlter = ds.Tables(0).Rows(intI)("AlterString")
                    dtExecute = Now()
                    command = RPT.GetSqlStringCommand(strAlter)
                    command.CommandTimeout = 0
                    RPT.ExecuteNonQuery(command)
                    command.Connection.Close()
                    dtComplete = Now()
                    elapsedSeconds = DateDiff(DateInterval.Second, dtExecute, dtComplete)
                    System.Threading.Thread.Sleep(elapsedSeconds * 1000)
                Next
            End If
            ds.Dispose()
        Catch ex As Exception
            lstMsgs(ex.Message & " for " & strTable)
        Finally
            lstMsgs("Finished in " & DateDiff(DateInterval.Minute, startTime, Now) & " minute(s).")
        End Try
        lstMsgs(".")
    End Sub
#End Region

#Region "Timer Manipulations"
    Const IntervalInMiliSecond As Integer = 10000

    Private Sub ResumeTimer(ByVal myTimer As System.Threading.Timer)
        myTimer.Change(IntervalInMiliSecond, IntervalInMiliSecond)
    End Sub

    Private Sub PauseTimer(ByVal myTimer As System.Threading.Timer)
        myTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
    End Sub

    Function GetFirstInterval() As Int64
        Return (DateDiff(DateInterval.Second, Now, CDate(Now.AddHours(1).ToString("yyyy-MM-dd HH:00")))) * 1000
    End Function

    Sub ResetTimer(ByVal myTimer As System.Threading.Timer)
        myTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
        myTimer.Change(IntervalInMiliSecond, IntervalInMiliSecond)
    End Sub
#End Region

#Region "GPS Location"
    Public Const DELIMITOR_COLON As String = ":"

    Private Function ParseData(ByVal _data As String) As AVLProperties
        Dim _errorId, _errorMsg As String
        Dim splitData As String()

        Dim avl As New AVLProperties

        Try
            splitData = _data.Split(DELIMITOR_COLON)

            If (splitData Is Nothing) Then
                _errorId = "ERR0008"
                _errorMsg = "[Data] is NULL or Empty"
                'Return ReturnStatus.CI_Fail
            End If

            If (splitData.Length <= 0) Then
                _errorId = "ERR0009"
                _errorMsg = "[Data] length is 0"
                'Return ReturnStatus.CI_Fail
            End If

            avl._enable = splitData(4)
            avl._alarm = splitData(5)
            avl._speed = splitData(6)
            avl._heading = splitData(8) + splitData(7)
            avl._longitude = splitData(12) + splitData(11) + splitData(10) + splitData(9)
            avl._latitude = splitData(16) + splitData(15) + splitData(14) + splitData(13)
            avl._strGPSDatetime = splitData(20) + splitData(19) + splitData(18) + splitData(17)
            avl._trackerId = splitData(21) + splitData(22) + splitData(23) + splitData(24) + _
            splitData(25) + splitData(26) + splitData(27) + splitData(28) + splitData(29) + splitData(30) + splitData(31)

            'Return ReturnStatus.CI_OK
        Catch ex As Exception
            _errorId = "ERR0010"
            _errorMsg = ex.Message
            'Return ReturnStatus.CI_Fail
        Finally
            lstMsgs(_errorId & vbTab & _errorMsg)
        End Try
    End Function
#End Region

#Region "GPS Datetime"

    Private Sub CheckAVLDatetime()

        Dim TrackerID As String = "NR09G20844"
        ReCaptionForm(TrackerID)
        'CLEARLIST = False

        Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("AVL")

        Dim Sql As New System.Text.StringBuilder

        Sql.Append(" SELECT * ")
        Sql.Append(" FROM [avldb].[dbo].[avl_raw] where [Tracker_ID] = '" & TrackerID & "' ")
        Sql.Append(" And DatePart(Hour, [Received_Datetime]) = 17")
        'Sql.Append(" and DATEPART(MINUTE, [Received_Datetime])=0")
        Sql.Append(" order by [Received_Datetime]")

        Dim Command As System.Data.Common.DbCommand = db.GetSqlStringCommand(Sql.ToString)
        Dim dt As DataTable = db.ExecuteDataSet(Command).Tables(0)
        Command.Connection.Close()

        Debug.Print(dt.Rows.Count)

        For Each nRow As DataRow In dt.Rows
            lstMsgs(GPSDatetime(nRow("Raw_Data")) & vbTab & nRow("Received_Datetime") & vbTab & GetTrackerID(nRow("Raw_Data")))
        Next

        dt.Dispose()
        Command.Dispose()

        GC.Collect()

    End Sub

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

    Private Sub CheckAVLFrequencies()

        Dim TrackerID As String = "NR09G20096"
        ReCaptionForm(TrackerID)
        'CLEARLIST = False

        Dim db As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase("AVL")

        Dim Sql As New System.Text.StringBuilder

        Sql.Append(" SELECT [Tracker_ID],[Received_Datetime],count(1)bil ")
        Sql.Append(" FROM [avldb].[dbo].[avl_raw] where [Tracker_ID] = '" & TrackerID & "' ")
        Sql.Append(" And DatePart(Hour, [Received_Datetime]) = 19")
        Sql.Append(" GROUP BY [Tracker_ID],[Received_Datetime]")
        Sql.Append(" order by [Received_Datetime]")

        Dim Command As System.Data.Common.DbCommand = db.GetSqlStringCommand(Sql.ToString)
        Dim dt As DataTable = db.ExecuteDataSet(Command).Tables(0)
        Command.Connection.Close()

        For Each nRow As DataRow In dt.Rows
            lstMsgs(nRow("Tracker_ID") & vbTab & nRow("Received_Datetime") & vbTab & nRow("bil"))
        Next

        dt.Dispose()
        Command.Dispose()

        GC.Collect()

    End Sub

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

#Region "Web Screenshot"

    Private Sub TestScreen()
        Dim url As String = "http://211.25.170.211/jksb/"
        Dim thumbnail As Bitmap = GenerateScreenshot(url, 800, 800)
        thumbnail.Save("test.png", System.Drawing.Imaging.ImageFormat.Png)
    End Sub

    Private Function GenerateScreenshot(ByVal url As String, ByVal width As Integer, ByVal height As Integer) As Bitmap

        Dim uid As String = "JKSBADMIN"
        Dim pwd As String = "12345"

        'Load the webpage into a WebBrowser control
        Dim wb As WebBrowser = New WebBrowser()

        wb.ScrollBarsEnabled = False
        wb.ScriptErrorsSuppressed = True
        wb.Navigate(url)
        While (wb.ReadyState <> WebBrowserReadyState.Complete)
            Application.DoEvents()
        End While

        Try

            wb.Document.GetElementById("ctl00_mainMasterContent_txtUsername").SetAttribute("value", uid)
            wb.Document.GetElementById("ctl00_mainMasterContent_txtPassword").SetAttribute("value", pwd)
            'Dim login As HtmlElement = wb.Document.GetElementById("ctl00_mainMasterContent_btnLogin")

            wb.AllowNavigation = True
            wb.Document.Forms(0).InvokeMember("submit")
            'wb.Document.Forms.GetElementsByName("aspnetForm").Item(0).InvokeMember("submit")

            wb.Navigate("http://211.25.170.211/jksb/DRS_Overview_JKSB.aspx")
            While (wb.ReadyState <> WebBrowserReadyState.Complete)
                Application.DoEvents()
            End While

        Catch ex As Exception
            lstMsgs(ex.Message)
        End Try

        'Set the size of the WebBrowser control
        wb.Width = width
        wb.Height = height

        If (width = -1) Then
            'Take Screenshot of the web pages full width
            wb.Width = wb.Document.Body.ScrollRectangle.Width
        End If

        If (height = -1) Then
            'Take Screenshot of the web pages full height
            wb.Height = wb.Document.Body.ScrollRectangle.Height
        End If

        If wb.DocumentText.Contains("Runtime Error") Then
            Return New Bitmap(width, height)
        End If

        'Get a Bitmap representation of the webpage as it's rendered in the WebBrowser control
        Dim bitmap As Bitmap = New Bitmap(wb.Width, wb.Height)
        wb.DrawToBitmap(bitmap, New Rectangle(0, 0, wb.Width, wb.Height))
        wb.Dispose()

        Return bitmap
        'Return crop(bitmap)

    End Function

    Private Function crop(ByVal OriginalImage As Bitmap) As Bitmap
        Dim CropRect As New Rectangle(5, 5, 530, 290)
        Dim CropImage = New Bitmap(CropRect.Width, CropRect.Height)
        Using grp = Graphics.FromImage(CropImage)
            grp.DrawImage(OriginalImage, New Rectangle(0, 0, CropRect.Width, CropRect.Height), CropRect, GraphicsUnit.Pixel)
            OriginalImage.Dispose()
            Return CropImage
        End Using
    End Function

    Private Function rescale(ByVal OriginalImage As Bitmap, ByVal ScaleFactor As Single) As Bitmap
        Try
            ' Get the scale factor.
            Dim scale_factor As Single = Single.Parse(ScaleFactor)

            ' Get the source bitmap.
            Dim bm_source As New Bitmap(OriginalImage)

            ' Make a bitmap for the result.
            Dim bm_dest As New Bitmap( _
                CInt(bm_source.Width * scale_factor), _
                CInt(bm_source.Height * scale_factor))

            ' Make a Graphics object for the result Bitmap.
            Dim gr_dest As Graphics = Graphics.FromImage(bm_dest)

            ' Copy the source image into the destination bitmap.
            gr_dest.DrawImage(bm_source, 0, 0, _
                bm_dest.Width + 1, _
                bm_dest.Height + 1)

            ' Display the result.
            Return bm_dest
        Catch ex As Exception
            Return OriginalImage 'On Exception return original image
        End Try
    End Function

#End Region

#End Region

End Class

Public Class Test
    Public Sub Execute(ByRef str As String)
        subtest(str)
    End Sub
    Private Sub subtest(ByRef x As String)
        x = "DAMN"
    End Sub
End Class