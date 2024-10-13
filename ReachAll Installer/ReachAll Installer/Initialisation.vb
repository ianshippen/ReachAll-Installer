Imports System.Xml

Module Initialisation
    Public reachallApplicationFileList As New List(Of String)

    Const CONFIG_FILENAME As String = "ReachAll Installer Config.xml"

    Public Sub Initialise()
        InitConfig()
        LoadConfig()
    End Sub

    Private Sub InitConfig()
        reachallApplicationFileList.Clear()
    End Sub

    Private Sub LoadConfig()
        Dim myFilename As String = GetApplicationPath().CreateFullName(CONFIG_FILENAME)

        ' Check that the file exists before trying to load
        Dim myFileInfo As New IO.FileInfo(myFilename)

        If myFileInfo.Exists Then
            Dim myDoc As New XmlDocument
            Dim myRecord As XmlNode = Nothing

            myDoc.Load(myFilename)

            ' Loop over each parameter
            For Each myRecord In myDoc("Config")
                Select Case myRecord.Name
                    Case "ReachAllApplicationSettings"
                        For Each a As XmlNode In myRecord
                            Select Case a.Name
                                Case "Files"
                                    For Each b As XmlNode In a.ChildNodes
                                        Select Case b.Name
                                            Case "File"
                                                If b.HasChildNodes Then reachallApplicationFileList.Add(b.FirstChild.Value)
                                        End Select
                                    Next
                            End Select
                        Next

                End Select
            Next
        Else
            '            Logutil.LogError("Config.vb", "Cannot find config file: " & myFilename)
        End If
    End Sub

End Module
