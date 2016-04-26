Imports System.Drawing

Public Class Form1

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
        'If (DataGridView1.InvokeRequired) Then
        '    DataGridView1.Invoke( _
        '            New PopulateGridViewDelegate(AddressOf PopulateGridView), _
        '            New Object() {DataGridView1, nLog})
        'Else
        '    DataGridView1.DataSource = nLog
        '    If nLog.Rows.Count > 0 Then
        '        DataGridView1.AutoResizeColumns()
        '        DataGridView1.PerformLayout()
        '        DataGridView1.Refresh()
        '    End If
        'End If
    End Sub

    Private Sub UpdateStatus(ByVal AddText As String)
        If (StatusStrip1.InvokeRequired) Then
            StatusStrip1.Invoke(New updateStatusDelegate(AddressOf AddItemsToStatus), New Object() {StatusStrip1, CStr(AddText)})
        Else
            Me.StatusStrip1.Items.Clear()
            Me.StatusStrip1.Items.Add(AddText)
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "fman"
        Dim nSize As Size = Me.Size
        Me.MaximumSize = nSize

        With txtPath
            .Text = "C:\download\drive-u\img\whatsapp\test"
            .ReadOnly = True
        End With

        With txtPattern
            .Text = "IMG-yyyymmdd-WA###"
            .ReadOnly = True
        End With

        AddHandler Me.Resize, AddressOf formResize

        Steps()

    End Sub

    Private Sub Steps()
        Dim startTime As Date = Now
        Try
            Dim invalidDate As Date = Nothing
            For Each x In GetFileList(txtPath.Text, "*.jpg")

                Dim nDate As Date = getDateFromName(x.Name)
                If Not nDate = invalidDate Then
                    Dim newPath As String = System.IO.Path.Combine(txtPath.Text, constructPathString(nDate))
                    If isValidPath(newPath) = False Then CreatePath(newPath)
                    System.IO.File.Move(System.IO.Path.Combine(txtPath.Text, x.Name), System.IO.Path.Combine(newPath, x.Name))
                    lstMsgs(x.Name & " moved.")
                Else
                    lstMsgs(x.Name & " failed to move.")
                End If

            Next
        Catch ex As Exception
            lstMsgs(ex.Message)
        End Try

        lstMsgs("Process finished in " & getTimeString(DateDiff(DateInterval.Second, startTime, Now)))

    End Sub

    Private Function getDateFromName(ByVal strName As String) As Date
        Dim result As Date = Nothing
        Try
            Dim tmp As String = strName.Remove(0, 4).Split("-")(0)
            Dim y, m, d
            y = CInt(tmp.Substring(0, 4))
            m = CInt(tmp.Substring(4, 2))
            d = CInt(tmp.Substring(6, 2))

            result = New Date(y, m, d)

        Catch ex As Exception
            result = Nothing
        End Try
        Return result
    End Function

    Private Function constructPathString(ByVal nDate As Date) As String
        Return nDate.Year & "\" & MonthName(nDate.Month, True).ToUpper
    End Function

    Private Sub formResize(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Top_Height As Double = 0.5
        Dim Bottom_Height As Double = 0.5
        Dim Width_Margin As Double = IIf(My.Computer.Info.OSVersion.StartsWith("5"), 3, 5) 'WIN7++ ADJUSTMENT

        Dim Left_Width As Double = 0.5
        Dim Right_Width As Double = 0.5

        'With GroupData
        '    .Location = New System.Drawing.Point(Width_Margin, 0)
        '    .Height = Me.Height * Top_Height
        '    .Width = Me.Width - (5 * Width_Margin)
        'End With

        With GroupVerbose
            .Location = New System.Drawing.Point(Width_Margin, Me.Height * Top_Height)
            .Width = Me.Width - (5 * Width_Margin)
            .Height = (Me.Height * Bottom_Height) - (30 + StatusStrip1.Height)
        End With

    End Sub

#Region "Methods"
    Public Function GetFileList(ByVal DIR As String, ByVal EXT As String) As IO.FileInfo()
        Dim di As New IO.DirectoryInfo(DIR)
        Return di.GetFiles(EXT)
    End Function

    Public Function isValidPath(ByVal filepath As String) As Boolean
        If System.IO.File.Exists(filepath) = True Then
            Return True
        ElseIf (IO.Directory.Exists(filepath) = True) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub CreatePath(ByVal path As String)
        If Not System.IO.Directory.Exists(path) Then
            System.IO.Directory.CreateDirectory(path)
        End If
    End Sub

    Private Function getTimeString(ByVal seconds As Integer) As String
        Dim t As TimeSpan = TimeSpan.FromSeconds(seconds)
        'Return String.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms", t.Hours, t.Minutes, t.Seconds, t.Milliseconds)
        Return String.Format("{0:D2}h:{1:D2}m:{2:D2}s", t.Hours, t.Minutes, t.Seconds)
    End Function

#End Region

End Class
