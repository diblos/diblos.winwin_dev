Public Class Form2

    Dim nSize As Size
    Dim _datatable As DataTable
    Dim WithEvents _timer As System.Timers.Timer
    Dim _srv As AVL_Raw_Srv

#Region "Form Events"
    Public Delegate Sub PopulateGridViewDelegate( _
                        ByVal GridView As DataGridView, _
                        ByVal nLog As DataTable)

    Public Delegate Sub AddItemsToListBoxDelegate( _
                         ByVal ToListBox As ListBox, _
                         ByVal AddText As String)

    Private Sub AddItemsToListBox(ByVal ToListBox As ListBox, _
                                 ByVal AddText As String)
        If ToListBox.Items.Count > 1000 Then
            ToListBox.Items.Clear()
        End If

        ToListBox.Items.Add(AddText)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, True)
        ToListBox.SetSelected(ListBox1.Items.Count - 1, False)
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
#End Region

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not IsNothing(_srv) Then _srv = Nothing
        If Not IsNothing(_datatable) Then _datatable.Dispose()
        'If Not IsNothing(_timer) Then _timer.
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        nSize = Me.Size
        Me.MaximumSize = nSize

        Me.Text = "Test Datatable"
        DataGridView1.AllowUserToAddRows = False

        'Dim _thread As New System.Threading.Thread(AddressOf Test1)
        '_thread.Start()

        _srv = New AVL_Raw_Srv
        _datatable = _srv.getInitialData

        _timer = New System.Timers.Timer()
        AddHandler _timer.Elapsed, AddressOf Test1
        _timer.Start()

    End Sub

    Private Sub Test1(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)

        Dim tmr As System.Timers.Timer = DirectCast(sender, System.Timers.Timer)

        With tmr
            .Stop()
            lstMsgs(updateDT(Now.Second, Now).ToString)
            ShowData(_datatable)
            If .Interval <> 1000 Then .Interval = 1000
            .Start()
        End With

    End Sub

    Private Function updateDT(ByVal ID As String, ByVal Timestamp As Date) As IncomingStatus
        Dim row() As DataRow = _datatable.Select("TrackerID='" & ID & "'")
        If row.Count > 0 Then
            If DateDiff(DateInterval.Second, row(0)("GPSDatetime"), Timestamp) < 61 Then
                Return IncomingStatus.Ignored
            Else
                row(0)("GPSDatetime") = Timestamp
                Return IncomingStatus.Updated
            End If
        Else
            Dim _row As DataRow
            _row = _datatable.NewRow
            _row("TrackerID") = ID
            _row("GPSDatetime") = Timestamp
            _datatable.Rows.Add(_row)
            Return IncomingStatus.Inserted
        End If
    End Function

    Private Enum IncomingStatus
        Inserted
        Updated
        Ignored
    End Enum
End Class