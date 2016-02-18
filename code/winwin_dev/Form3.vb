Imports Itramas.Itac.Helper
Imports System.Net.NetworkInformation

Public Class Form3

    Dim nSize As New Size(600, 400)
    Dim CLEARLIST As Boolean = True
    Dim dtTableSize As DataTable

    Dim DBLOGFILE_FLAG As Boolean = False

    Dim counter As Integer = 0

#Region "Form Delegates"
    Public Delegate Sub AddItemsToListBoxDelegate( _
                  ByVal ToListBox As ListBox, _
                  ByVal AddText As String)

    Public Delegate Sub updateItemsToListBoxDelegate( _
                 ByVal ToListBox As ListBox, _
                 ByVal AddText As String)

    Public Delegate Sub updateStatusDelegate( _
             ByVal ToListBox As StatusStrip, _
             ByVal AddText As String)

    Public Delegate Sub disableFormDelegate( _
            ByVal mForm As Form3, _
            ByVal bool As Boolean)

    Public Delegate Sub UpdateDataGridDelegate(ByVal DataGridView As DataGridView, ByVal DataSource As Object)

    Private Sub AddItemsToListBox(ByVal ToListBox As ListBox, _
                                 ByVal AddText As String)
        If ToListBox.Items.Count > 1000 And CLEARLIST = True Then
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
            If Me.ListBox1.Items.Count > 1000 And CLEARLIST = True Then
                ListBox1.Items.Clear()
            End If

            Me.ListBox1.Items.Add(CStr(item))
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, True)
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, False)
        End If

    End Sub

    Private Sub AddItemsToStatus(ByVal StatusStrips As StatusStrip, _
                             ByVal AddText As String)
        StatusStrips.Items.Clear()
        StatusStrips.Items.Add(AddText)
    End Sub

    Private Sub updateListBox(ByVal ToListBox As ListBox, _
                             ByVal AddText As String)
        ToListBox.Items(ToListBox.Items.Count - 1) = ToListBox.Items(ToListBox.Items.Count - 1) & " " & AddText
        ToListBox.SetSelected(ListBox1.Items.Count - 1, True)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, False)
    End Sub

    Public Sub UpdateMsgs(ByVal item As Object, Optional ByVal source As Object = Nothing)
        If (ListBox1.InvokeRequired) Then
            ListBox1.Invoke( _
                    New updateItemsToListBoxDelegate(AddressOf updateListBox), _
                    New Object() {ListBox1, CStr(item)})
        Else
            Me.ListBox1.Items(Me.ListBox1.Items.Count - 1) = Me.ListBox1.Items(Me.ListBox1.Items.Count - 1) & " " & item
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, True)
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, False)
        End If

    End Sub

    Private Sub UpdateStatus(ByVal AddText As String)
        If (StatusStrip1.InvokeRequired) Then
            StatusStrip1.Invoke( _
                    New updateStatusDelegate(AddressOf AddItemsToStatus), _
                    New Object() {StatusStrip1, CStr(AddText)})
        Else
            Me.StatusStrip1.Items.Clear()
            Me.StatusStrip1.Items.Add(AddText)
        End If
    End Sub

    Private Sub disableFormElements(ByVal mForm As Form3, _
                             ByVal bool As Boolean)

        mForm.Button1.Enabled = bool
        If DBLOGFILE_FLAG = False Then mForm.ButtonPurge.Enabled = bool
        mForm.ListTable.Enabled = bool
        mForm.DataGridView1.Enabled = bool

        If bool Then
            mForm.Cursor = Cursors.Default
        Else
            mForm.Cursor = Cursors.WaitCursor
        End If
    End Sub

    Private Sub DisableForm(ByVal bool As Boolean)

        If (Me.InvokeRequired) Then
            Me.Invoke( _
                    New disableFormDelegate(AddressOf disableFormElements), _
                    New Object() {Me, bool})
        Else

            Me.Button1.Enabled = bool
            If DBLOGFILE_FLAG = False Then Me.ButtonPurge.Enabled = bool
            Me.ListTable.Enabled = bool
            Me.DataGridView1.Enabled = bool

            If bool Then
                Me.Cursor = Cursors.Default
            Else
                Me.Cursor = Cursors.WaitCursor
            End If
        End If

    End Sub

    Private Sub updateDGV(ByVal dgv As DataGridView, _
                          ByVal DataSource As Object)
        dgv.DataSource = DataSource
        dgv.AutoResizeColumns()
    End Sub

    Private Sub UpdateDataGrid(ByVal DataSource As Object)

        If (DataGridView1.InvokeRequired) Then
            DataGridView1.Invoke( _
                    New UpdateDataGridDelegate(AddressOf updateDGV), _
                    New Object() {DataGridView1, DataSource})
        Else
            DataGridView1.DataSource = DataSource
            DataGridView1.AutoResizeColumns()
        End If

    End Sub

#End Region

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If IsNothing(dtTableSize) = False Then dtTableSize.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            End
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim _encrypt As New Encryption
        With _encrypt
            '.ProtectApplicationSetting()
            '.ProtectConnectionString()
        End With

        Initialization()
        Me.Text = APPLICATION_NAME
        Me.Size = nSize
        Me.MinimumSize = nSize
        Me.MaximumSize = nSize

        Dim minToTray As New clsMinToTray(Me, APPLICATION_NAME & " is running...", Me.Icon)

        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
        DataGridView1.AllowUserToAddRows = False

        GetTables()
        GetTableSizes()

        UpdateStatus("Ready")

        'Debug.Print("OS         : " & My.Computer.Info.OSFullName & vbNewLine & _
        '            "Version    : " & My.Computer.Info.OSVersion & vbNewLine & _
        '            "Platform   : " & My.Computer.Info.OSPlatform & vbNewLine)

        'lstMsgs("OS         : " & My.Computer.Info.OSFullName)
        'lstMsgs("Version    : " & My.Computer.Info.OSVersion)
        'lstMsgs("Platform   : " & My.Computer.Info.OSPlatform)

        AddHandler ListTable.SelectedIndexChanged, AddressOf ListTable_SelectedIndexChanged
        AddHandler ButtonPurge.Click, AddressOf ButtonPurge_Click
        AddHandler Button1.Click, AddressOf Button1_Click

    End Sub

    Private Function getTimeString(ByVal seconds As Integer) As String
        Dim t As TimeSpan = TimeSpan.FromSeconds(seconds)
        'Return String.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms", t.Hours, t.Minutes, t.Seconds, t.Milliseconds)
        Return String.Format("{0:D2}h:{1:D2}m:{2:D2}s", t.Hours, t.Minutes, t.Seconds)
    End Function

#Region "Rebuild Index"

    Private Sub GetTables()

        Dim RPT As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
       Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase(SELECTED_CONNECTION)

        lstMsgs("Get " & GetCatalogDB(RPT.ConnectionString) & " tables.")

        Dim sql As String
        Dim ds As DataSet
        Dim command As System.Data.Common.DbCommand

        Try
            sql = _
            " SELECT TABLE_NAME FROM information_schema.tables WHERE TABLE_TYPE='BASE TABLE' " & _
            " AND NOT (TABLE_NAME LIKE 'sys%' OR TABLE_NAME LIKE 'temp%' OR TABLE_NAME LIKE '%_2014%') " & _
            " ORDER BY TABLE_NAME "

            command = RPT.GetSqlStringCommand(sql)
            command.CommandTimeout = 0
            ds = RPT.ExecuteDataSet(command)
            command.Connection.Close()

            For Each row As DataRow In ds.Tables(0).Rows
                ListTable.Items.Add(row(0))
            Next

            command.Dispose()
            ds.Dispose()
        Catch ex As Exception
            lstMsgs(ex.Message & " for " & SELECTED_CONNECTION)
        Finally
            If Not IsNothing(RPT) Then RPT = Nothing
        End Try

    End Sub

    Private Sub GetTableSizes()
        Dim RPT As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
       Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase(SELECTED_CONNECTION)

        Dim sql As New System.Text.StringBuilder
        Dim command As System.Data.Common.DbCommand

        Try
            sql.Append(" CREATE TABLE #RowCountsAndSizes (TableName NVARCHAR(128),rows CHAR(11),     ")
            sql.Append(" reserved VARCHAR(18),data VARCHAR(18),index_size VARCHAR(18), ")
            sql.Append("       unused VARCHAR(18)) ")
            sql.AppendLine("       EXEC sp_MSForEachTable 'INSERT INTO #RowCountsAndSizes EXEC sp_spaceused ''?'' ' ")
            sql.AppendLine(" SELECT     TableName,CONVERT(bigint,rows) AS NumberOfRows, ")
            sql.Append("           CONVERT(bigint,left(reserved,len(reserved)-3)) AS SizeinKB ")
            sql.Append(" FROM       #RowCountsAndSizes ")
            sql.Append(" ORDER BY   NumberOfRows DESC,SizeinKB DESC,TableName ")
            sql.AppendLine(" DROP TABLE #RowCountsAndSizes ")

            command = RPT.GetSqlStringCommand(sql.ToString)
            command.CommandTimeout = 0
            dtTableSize = RPT.ExecuteDataSet(command).Tables(0)
            command.Connection.Close()
            command.Dispose()
        Catch ex As Exception
            dtTableSize = New DataTable
            lstMsgs("GetTableSizes: " & ex.Message)
        Finally
            If Not IsNothing(sql) Then sql = Nothing
            If Not IsNothing(RPT) Then RPT = Nothing
        End Try
    End Sub

    Private Sub QueryIndexFragmenttations(ByVal tablename As String)
        lstMsgs("query fragmentation for table " & tablename & "...")
        Dim RPT As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase(SELECTED_CONNECTION)

        Dim sql As New System.Text.StringBuilder
        Dim ds As DataSet
        Dim command As System.Data.Common.DbCommand

        Try
            sql.Append(" SELECT index_id [index],ROUND(avg_fragmentation_in_percent, 5) avg_fragmentation_in_percent, fragment_count ")
            sql.Append(" FROM sys.dm_db_index_physical_stats (DB_ID(), OBJECT_ID('" & tablename & "'), NULL, NULL, 'LIMITED') ")
            sql.Append(" GO ")

            command = RPT.GetSqlStringCommand(sql.ToString)
            command.CommandTimeout = 0
            ds = RPT.ExecuteDataSet(command)
            command.Connection.Close()

            UpdateDataGrid(ds.Tables(0))

            command.Dispose()
            ds.Dispose()
        Catch ex As Exception
            lstMsgs(ex.Message & " for " & SELECTED_CONNECTION)
        Finally
            If Not IsNothing(RPT) Then RPT = Nothing
            UpdateMsgs("Done.")
        End Try

    End Sub

    Private Sub RebuildTableIndex(ByVal strTable As String)
        DisableForm(False)
        Dim startTime As Date = Now
        lstMsgs("Rebuilding Table Index: " & strTable & "...")
        Dim RPT As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
       Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase(SELECTED_CONNECTION)

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
                    Application.DoEvents()
                Next
            End If

            command.Dispose()
            ds.Dispose()
        Catch ex As Exception
            lstMsgs(ex.Message & " for " & strTable)
        Finally
            If Not IsNothing(RPT) Then RPT = Nothing
            UpdateMsgs("Done in " & getTimeString(DateDiff(DateInterval.Second, startTime, Now)))
        End Try
        DisableForm(True)
        GC.Collect()
    End Sub
#End Region

#Region "Shrink Database Log File"

    Private Function GetCatalogDB(ByVal ConnectionStr As String) As String
        Dim tmp() As String = ConnectionStr.Split(";")
        Dim result As String = String.Empty
        For Each x As String In tmp
            If UCase(x).StartsWith("INITIAL CATALOG") Then
                result = x.Split("=")(1)
            End If
        Next
        Return result
    End Function

    Private Sub ShrinkDBLog()
        DisableForm(False)
        Dim startTime As Date = Now

        Dim RPT As Microsoft.Practices.EnterpriseLibrary.Data.Database = _
        Microsoft.Practices.EnterpriseLibrary.Data.DatabaseFactory.CreateDatabase(SELECTED_CONNECTION)

        Dim sql As New System.Text.StringBuilder
        Dim ds As DataSet
        Dim command As System.Data.Common.DbCommand

        Dim strDBName As String = GetCatalogDB(RPT.ConnectionString)
        lstMsgs("Shrinking DB Log: " & strDBName & "...")

        Try
            sql.AppendLine(" USE @DBNAME ;")
            sql.AppendLine(" ALTER DATABASE @DBNAME ")
            sql.AppendLine(" SET RECOVERY SIMPLE ;")
            sql.AppendLine(" DBCC SHRINKFILE (@DBLOG, 1) ;")
            sql.AppendLine(" ALTER DATABASE @DBNAME ")
            sql.AppendLine(" SET RECOVERY FULL ;")

            sql.Replace("@DBNAME", strDBName)
            sql.Replace("@DBLOG", strDBName & "_log")
            command = RPT.GetSqlStringCommand(sql.ToString)

            command.CommandTimeout = 0
            ds = RPT.ExecuteDataSet(command)
            command.Connection.Close()
            command.Dispose()
            'USE DS DATA OR ...
            ds.Dispose()
        Catch ex As Exception
            lstMsgs(ex.Message & " for " & strDBName)
        Finally
            If Not IsNothing(RPT) Then RPT = Nothing
            UpdateMsgs("Done in " & getTimeString(DateDiff(DateInterval.Second, startTime, Now)))
        End Try
        DisableForm(True)
        GC.Collect()
    End Sub
#End Region

    Private Sub Form3_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Dim Top_Height As Double = 0.5
        Dim Bottom_Height As Double = 0.5
        Dim Width_Margin As Double = IIf(My.Computer.Info.OSVersion.StartsWith("5"), 3, 5) 'WIN7++ ADJUSTMENT

        Group1.Location = New Point(Width_Margin, 0)
        Group1.Height = Me.Height * Top_Height - ButtonPurge.Height - 10
        Group1.Width = (Me.Width / 2) - (2 * Width_Margin)

        Group2.Location = New Point((Me.Width / 2), 0)
        Group2.Height = Me.Height * Top_Height - Button1.Height - 10
        Group2.Width = (Me.Width / 2) - (4 * Width_Margin)

        GroupVerbose.Location = New Point(Width_Margin, Me.Height * Bottom_Height)
        GroupVerbose.Width = Me.Width - (5 * Width_Margin)
        GroupVerbose.Height = (Me.Height * Bottom_Height) - (30 + StatusStrip1.Height)

        'DataGridView1.Height = Button1.Top - Button1.Height

        ButtonPurge.Width = Group1.Width
        ButtonPurge.Location = New Point(Group1.Location.X, Group1.Location.Y + Group1.Height + 5)

        Button1.Width = Group2.Width
        Button1.Location = New Point(Group2.Location.X, Group2.Location.Y + Group2.Height + 5)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ListTable.SelectedIndex <> -1 Then
            Dim k As New System.Threading.Thread(AddressOf RebuildTableIndex)
            k.Start(ListTable.Items(ListTable.SelectedIndex))
        End If
    End Sub

    Private Sub ListTable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim dListBox As ListBox = DirectCast(sender, ListBox)
        Dim str As String = dListBox.Items(dListBox.SelectedIndex)
        If str <> "" Then
            Dim q As New System.Threading.Thread(AddressOf QueryTableInfo)
            q.Start(str)
        End If

    End Sub

    Private Sub QueryTableInfo(ByVal tablename As String)
        DisableForm(False)
        QueryIndexFragmenttations(tablename)
        readSizeDT(tablename)
        DisableForm(True)
        GC.Collect()
    End Sub

    Private Sub readSizeDT(ByVal TableName As String)
        If IsNothing(dtTableSize) Then StatusStrip1.Text = "Table: " & TableName & " - info Not available."
        Dim row() As DataRow = dtTableSize.Select("TableName='" & TableName & "'")
        If row.Count > 0 Then
            UpdateStatus("Table: " & TableName & ", count: " & row(0)("NumberOfRows") & " row(s), size: " & row(0)("SizeinKB") & " KB")
        Else
            UpdateStatus("Table: " & TableName & " - info Not available.")
        End If
    End Sub

    Private Sub ButtonPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim t As New Threading.Thread(AddressOf ShrinkDBLog)
        t.Start()
        DBLOGFILE_FLAG = True
        ButtonPurge.Enabled = False
    End Sub
End Class
