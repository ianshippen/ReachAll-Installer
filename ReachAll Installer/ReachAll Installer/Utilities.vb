Module Utilities
    Public Function GetApplicationPath() As String
        Dim result As String = System.AppDomain.CurrentDomain.BaseDirectory()

        Return result
    End Function
End Module
