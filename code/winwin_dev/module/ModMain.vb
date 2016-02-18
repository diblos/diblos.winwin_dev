Module ModMain

    Public Const DEFAULT_DATESTRING As String = "yyyy-MM-dd HH:mm:ss"

    Public SELECTED_CONNECTION As String = String.Empty
    Public APPLICATION_NAME As String = String.Empty
    Public COMMAND_QUEUE As String

    Public Const DATETIME_FORMAT = "yyyy-MM-dd HH:mm:ss"
    Public dec0F As Integer = Convert.ToInt32("0x0F", 16)
    Public dec1F As Integer = Convert.ToInt32("0x1F", 16)
    Public dec3F As Integer = Convert.ToInt32("0x3F", 16)

    Public Sub Initialization()
        SELECTED_CONNECTION = IIf(Configuration.ConfigurationManager.AppSettings("SELECTED_CONNECTION") = "", "DB", Configuration.ConfigurationManager.AppSettings("SELECTED_CONNECTION"))
        APPLICATION_NAME = IIf(Configuration.ConfigurationManager.AppSettings("APPNAME") = "", "", Configuration.ConfigurationManager.AppSettings("APPNAME"))
        COMMAND_QUEUE = IIf(Configuration.ConfigurationManager.AppSettings("DRS_QUEUE") = "", "", Configuration.ConfigurationManager.AppSettings("DRS_QUEUE"))
    End Sub

End Module
