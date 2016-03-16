Imports System.Windows.Forms
Imports System.Threading
Imports System.ServiceProcess
Imports System.IO

Public Class frmWatcher

#Region "Declarations"
    Public watchfolder() As FileSystemWatcher

    Dim myIcon As New System.Drawing.Icon(fwatcher.My.Resources.Resources.Archigraphs_Cubed_Animals_Piggy, fwatcher.My.Resources.Resources.Archigraphs_Cubed_Animals_Piggy.Size)

    Dim nSize As New System.Drawing.Size(500, 400)

    Dim _datatable As DataTable

#End Region
#Region "Form / Development"

    Public Delegate Sub PopulateGridViewDelegate(ByVal GridView As DataGridView, ByVal nLog As DataTable)
    Public Delegate Sub AddItemsToListBoxDelegate(ByVal ToListBox As ListBox, ByVal AddText As String)
    Public Delegate Sub updateStatusDelegate(ByVal ToListBox As StatusStrip, ByVal AddText As String)

    Private Sub AddItemsToListBox(ByVal ToListBox As ListBox, _
                                 ByVal AddText As String)
        If ToListBox.Items.Count > 1000 Then
            ToListBox.Items.Clear()
        End If

        ToListBox.Items.Add(AddText)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, True)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, False)
    End Sub

    Public Sub lstMsgs(ByVal item As Object)
        If (ListBox1.InvokeRequired) Then
            ListBox1.Invoke( _
                    New AddItemsToListBoxDelegate(AddressOf AddItemsToListBox), _
                    New Object() {ListBox1, CStr(item)})
        Else
            If Me.ListBox1.Items.Count > 1000 Then
                ListBox1.Items.Clear()
            End If

            Me.ListBox1.Items.Add(CStr(item))
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, True)
            Me.ListBox1.SetSelected(ListBox1.Items.Count - 1, False)
        End If

    End Sub

    Private Sub AddItemsToStatus(ByVal StatusStrips As StatusStrip, ByVal AddText As String)
        StatusStrips.Items.Clear()
        StatusStrips.Items.Add(AddText)
    End Sub

    Private Sub PopulateGridView(ByVal GridView As DataGridView, _
                     ByVal nLog As DataTable)
        GridView.DataSource = nLog
        If nLog.Rows.Count > 0 Then
            GridView.AutoResizeColumns()
            GridView.PerformLayout()
            GridView.Refresh()
        End If
    End Sub

    Public Sub ShowData(ByVal nLog As Object)
        If (DataGridView1.InvokeRequired) Then
            DataGridView1.Invoke( _
                    New PopulateGridViewDelegate(AddressOf PopulateGridView), _
                    New Object() {DataGridView1, nLog})
        Else
            DataGridView1.DataSource = nLog
            If nLog.Rows.Count > 0 Then
                DataGridView1.AutoResizeColumns()
                DataGridView1.PerformLayout()
                DataGridView1.Refresh()
            End If
        End If

    End Sub

    Private Sub UpdateStatus(ByVal AddText As String)
        If (StatusStrip1.InvokeRequired) Then
            StatusStrip1.Invoke(New updateStatusDelegate(AddressOf AddItemsToStatus), New Object() {StatusStrip1, CStr(AddText)})
        Else
            Me.StatusStrip1.Items.Clear()
            Me.StatusStrip1.Items.Add(AddText)
        End If
    End Sub

#End Region
#Region "Functions"
    Private Sub setStatus(ByVal obj As Object)
        For x = 0 To NO_OF_APP - 1
            If PerfectPath(obj) = PerfectPath(APP_STATUS(x).WatchFolder) Then
                APP_STATUS(x).Status = AppStatus.Active
                APP_STATUS(x).Time = Now
            End If
        Next
    End Sub

    Private Sub logchange(ByVal source As Object, ByVal e As  _
                    System.IO.FileSystemEventArgs)
        If e.ChangeType = IO.WatcherChangeTypes.Changed Then
            If DEV_MODE = True Then
                lstMsgs(Now.ToString("yyyy-MM-dd HH:mm:ss") & " File " & e.FullPath & _
                        " has been modified" & vbCrLf)
            End If
        End If
        If e.ChangeType = IO.WatcherChangeTypes.Created Then
            If DEV_MODE = True Then
                lstMsgs(Now.ToString("yyyy-MM-dd HH:mm:ss") & " File " & e.FullPath & _
                        " has been created" & vbCrLf)
            End If
        End If
        If e.ChangeType = IO.WatcherChangeTypes.Deleted Then
            If DEV_MODE = True Then
                lstMsgs(Now.ToString("yyyy-MM-dd HH:mm:ss") & " File " & e.FullPath & _
                        " has been deleted" & vbCrLf)
            End If
        End If

        'lstMsgs(source.ToString)
        'lstMsgs(e.FullPath)
        'setStatus(GetWorkingDir(e.FullPath))

        UpdateStatus("Last modified: " & e.FullPath)
        updatesFileInfo()

    End Sub

    Public Sub logrename(ByVal source As Object, ByVal e As  _
                            System.IO.RenamedEventArgs)
        If DEV_MODE = True Then

            'If Me.InvokeRequired Then
            '    Me.Invoke(New MethodInvoker(AddressOf logrename))
            'Else
            '    lstMsgs("using textbox from another thread")
            'End If

            'lstResults.BeginInvoke((MethodInvoker)delegate() { lstResults.Columns.Clear(); }, null);

            lstMsgs(Now.ToString("yyyy-MM-dd HH:mm:ss") & " File" & e.OldName & _
                    " has been renamed to " & e.Name & vbCrLf)
        End If
    End Sub

    Private Sub StartWatching()

        Try
            ReDim watchfolder(NO_OF_APP - 1)

            For x = 0 To UBound(watchfolder)
                watchfolder(x) = New System.IO.FileSystemWatcher()

                'this is the path we want to monitor
                watchfolder(x).Path = APP_STATUS(x).WatchFolder

                'Add a list of Filter we want to specify

                'make sure you use OR for each Filter as we need to

                'all of those
                'watchfolder(x).NotifyFilter = IO.NotifyFilters.DirectoryName
                'watchfolder(x).NotifyFilter = watchfolder(x).NotifyFilter Or _
                '                              IO.NotifyFilters.FileName
                watchfolder(x).NotifyFilter = watchfolder(x).NotifyFilter Or _
                                              IO.NotifyFilters.Attributes
                ' add the handler to each event

                AddHandler watchfolder(x).Changed, AddressOf logchange
                AddHandler watchfolder(x).Created, AddressOf logchange
                AddHandler watchfolder(x).Deleted, AddressOf logchange

                ' add the rename handler as the signature is different
                'AddHandler watchfolder(x).Renamed, AddressOf logrename

                'Set this property to true to start watching

                watchfolder(x).EnableRaisingEvents = True
            Next

            'Button1.Text = "Stop Watch"
            'btn_startwatch.Enabled = False
            'btn_stop.Enabled = True

            'End of code for btn_start_click
        Catch ex As Exception
            lstMsgs(ex.Message)
        End Try

    End Sub

    Private Sub StopWatching()
        Try
            For x = 0 To UBound(watchfolder)
                watchfolder(x).EnableRaisingEvents = False
            Next
            'Button1.Text = "Start Watch"
            'btn_startwatch.Enabled = True
            'btn_stop.Enabled = False
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Public Function ProcessesRunning(ByVal ProcessName As String) As Integer
        '
        Try
            Return Process.GetProcessesByName(ProcessName).GetUpperBound(0) + 1
        Catch
            Return 0
        End Try

    End Function

    Private Sub formResize(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Top_Height As Double = 0.5
        Dim Bottom_Height As Double = 0.5
        Dim Width_Margin As Double = IIf(My.Computer.Info.OSVersion.StartsWith("5"), 3, 5) 'WIN7++ ADJUSTMENT

        Dim Left_Width As Double = 0.5
        Dim Right_Width As Double = 0.5

        With GroupData
            .Location = New System.Drawing.Point(Width_Margin, 0)
            .Height = Me.Height * Top_Height
            .Width = Me.Width - (5 * Width_Margin)
        End With

        With GroupVerbose
            .Location = New System.Drawing.Point(Width_Margin, Me.Height * Top_Height)
            .Width = Me.Width - (5 * Width_Margin)
            .Height = (Me.Height * Bottom_Height) - (30 + StatusStrip1.Height)
        End With

    End Sub

#End Region
#Region "Start / Stop"
    Private Sub frmWatcher_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        StopWatching()

        'If Not IsNothing(_srv) Then _srv = Nothing
        If Not IsNothing(_datatable) Then _datatable.Dispose()
    End Sub

    Private Sub frmWatcher_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If ProcessesRunning(myName) > 1 Then
            lstMsgs("Me is running (one instance)")
            End
        Else
            'INITIALIZE VALUES
            '--------------------------------------------------------
            Dim minToTray As New clsMinToTray(Me, myName & " is watching...", myIcon)
            InitAPP()
            _datatable = getInitialData()

            With DataGridView1
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .EditMode = DataGridViewEditMode.EditProgrammatically
            End With

            updatesFileInfo()

            'START WATCHER
            '--------------------------------------------------------
            StartWatching()

            'FORM ATTRIBUTES
            '--------------------------------------------------------
            AddHandler Me.Resize, AddressOf formResize
            Me.Icon = myIcon
            Me.Text = myName

            Me.Size = nSize
            Me.MaximumSize = nSize

            UpdateStatus("Ready")

        End If
    End Sub

    Private Function updateDT(ByVal ID As String, ByVal Timestamp As Date) As Object 'IncomingStatus
        Dim nRow() As DataRow = _datatable.Select("FileName='" & ID & "'")
        If nRow.Length > 0 Then
            'If DateDiff(DateInterval.Second, nRow(0)("LastUpdated"), Timestamp) < 61 Then
            '    Return Nothing 'IncomingStatus.Ignored
            'Else
            nRow(0)("LastUpdated") = Timestamp
            Return Nothing 'IncomingStatus.Updated
            'End If
        Else
            Dim _row As DataRow
            _row = _datatable.NewRow
            _row("No") = _datatable.Rows.Count + 1
            _row("FileName") = ID
            _row("LastUpdated") = Timestamp
            _datatable.Rows.Add(_row)
            Return Nothing 'IncomingStatus.Inserted
        End If
    End Function

    Private Sub updatesFileInfo()

        Try
            For Each x In GetFileList(APP_STATUS(0).WatchFolder, "*.xml")
                updateDT(x.Name, x.LastWriteTime)
            Next

            ShowData(_datatable)
            Debug.Print(_datatable.Rows.Count)
        Catch ex As Exception
            lstMsgs(ex.Message)
        End Try

    End Sub
#End Region

End Class