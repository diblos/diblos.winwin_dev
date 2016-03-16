Imports System.Configuration.ConfigurationManager

Module common
    Public Const DEVELOPMENT_MODE_KEY As String = "DEVELOPMENT.MODE"
    Public Const MINUTE_OF_INACTIVE_KEY As String = "MINUTE.OF.INACTIVE"
    Public Const DATE_FORMAT_DISPLAY As String = "yyyy-MM-dd HH:mm:ss"
    Public Const SEPARATOR As String = "|"
    Public myName As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
    Public RESTART_APPLICATION_PATH As String = "D:\workspace\_codes\POLRI\_SERVICE\AID_VDS\Itac.RestartApplication\bin\Debug\Itac.RestartApplication.exe"
    Public DEV_MODE As Boolean = False
    Public NO_OF_APP As Integer = AppSettings("NO.OF.APP")
    'Public INACTIVE_THRESHOLD_MINUTE As Integer = 1
    Public APP_STATUS() As ApplicationObject

    Public Sub InitAPP()
        Dim tmpStr As String
        ReDim APP_STATUS(NO_OF_APP - 1)

        If Not AppSettings(DEVELOPMENT_MODE_KEY) = "" Then
            Select Case UCase(AppSettings(DEVELOPMENT_MODE_KEY))
                Case "ON"
                    DEV_MODE = True
                Case "OFF"
                    DEV_MODE = False
                Case Else
                    DEV_MODE = False
            End Select
        End If

        For x = 0 To NO_OF_APP - 1

            tmpStr = AppSettings("APP." & x + 1)
            Dim arr As String = tmpStr '.Split(SEPARATOR)

            APP_STATUS(x) = New ApplicationObject

            'SET APPLICATION OBJECT VALUES
            '=============================
            'APP_STATUS(x).EXEPath = arr(1)
            APP_STATUS(x).Time = Now
            APP_STATUS(x).WatchFolder = arr
            APP_STATUS(x).Status = AppStatus.Active

            arr = Nothing
        Next
    End Sub

    Public Function getInitialData() As DataTable
        Dim nCol As DataColumn
        Dim tempData As New DataTable

        nCol = New DataColumn()
        nCol.DataType = System.Type.GetType("System.Int16")
        nCol.ColumnName = "No"
        nCol.ReadOnly = False
        nCol.Unique = False
        tempData.Columns.Add(nCol)

        nCol = New DataColumn()
        nCol.DataType = System.Type.GetType("System.String")
        nCol.ColumnName = "FileName"
        nCol.ReadOnly = False
        nCol.Unique = False
        tempData.Columns.Add(nCol)

        nCol = New DataColumn()
        nCol.DataType = System.Type.GetType("System.DateTime")
        nCol.ColumnName = "LastUpdated"
        nCol.ReadOnly = False
        nCol.Unique = False
        tempData.Columns.Add(nCol)

        Return tempData
    End Function

#Region "Respawn"
    'Same function as in frmTest
    Public Function RestartApp(ByVal App As ApplicationObject)
        Dim path As String = App.EXEPath
        Dim processName As String = RemoveFileExt(GetFileName(path))
        Dim result As Boolean = False
        Try
            For Each proc As Process In Process.GetProcesses
                'Kill process by  ProcessID
                If (proc.ProcessName = processName) Then
                    proc.Kill()
                End If
            Next

            Process.Start(path)
            result = True
        Catch ex As Exception
            WriteEventLog(EventLogEntryType.Error, "RestartApp: " & ex.ToString, myName, myName)
            result = False
            'Finally
        End Try
        RestartApp = result
    End Function
#End Region
End Module

Public Class ApplicationObject
    Private _Status As AppStatus
    Private _Time As DateTime
    Private _WatchFolder As String
    Private _Path As String

    Public Property Status() As AppStatus
        Get
            Return Me._Status
        End Get
        Set(ByVal value As AppStatus)
            Me._Status = value
        End Set
    End Property

    Public Property Time() As DateTime
        Get
            Return Me._Time
        End Get
        Set(ByVal value As DateTime)
            Me._Time = value
        End Set
    End Property

    Public Property WatchFolder() As String
        Get
            Return Me._WatchFolder
        End Get
        Set(ByVal value As String)
            Me._WatchFolder = value
        End Set
    End Property

    Public Property EXEPath() As String
        Get
            Return Me._Path
        End Get
        Set(ByVal value As String)
            Me._Path = value
        End Set
    End Property

End Class

Public Enum AppStatus
    Active = 1
    InActive = 2
End enum