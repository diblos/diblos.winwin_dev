Public Class AVLEngineClass
    '4 to 31 : C3:80:00:85:00:BC:74:CB:42:04:7E:47:40:14:40:9A:31:4E:52:30:39:47:30:33:31:32:34:00
    '22:00:08:00:C3:80:00:85:00:BC:74:CB:42:04:7E:47:40:14:40:9A:31:4E:52:30:39:47:30:33:31:32:34:00:00:00

    Dim DUMMY_STRING As String = "22:00:08:00:C3:80:00:85:00:BC:74:CB:42:04:7E:47:40:14:40:9A:31:4E:52:30:39:47:30:33:31:32:34:00:00:00"

    Public ReadOnly Property GetLength() As Integer
        Get
            Return DUMMY_STRING.Split(":").Length
        End Get
    End Property

    Public ReadOnly Property GetString() As String
        Get
            Return DUMMY_STRING
        End Get
    End Property

    Public Sub New()

    End Sub

End Class

Public Class AVLProperties
    Public _enable As String
    Public _alarm As String
    Public _speed As String
    Public _heading As String
    Public _longitude As String
    Public _latitude As String
    Public _strGPSDatetime As String
    Public _trackerId As String
End Class