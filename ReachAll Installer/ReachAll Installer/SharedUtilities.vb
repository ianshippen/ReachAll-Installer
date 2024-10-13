Imports IWshRuntimeLibrary
' Need to add reference to COM object Windows Script Host Object Model

Module SharedUtilities
    Public Function CreateShortCut(ByRef p_shortcutName As String, ByRef p_shortcutDirectory As String, ByRef p_targetPath As String, ByRef p_workingDirectory As String, ByRef p_iconFile As String, ByVal p_iconNumber As Integer) As Boolean
        Dim result As Boolean = False

        Try
            Dim myShell As New WshShell
            Dim shortCut As IWshRuntimeLibrary.IWshShortcut

            If Not IO.Directory.Exists(p_shortcutDirectory) Then
                If MsgBox("Shortcut folder " & p_shortcutDirectory & " does not exist. Do you want to create it ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    IO.Directory.CreateDirectory(p_shortcutDirectory)
                Else
                    Return False
                End If
            End If

            shortCut = CType(myShell.CreateShortcut(p_shortcutDirectory & "\" & p_shortcutName & ".lnk"), IWshRuntimeLibrary.IWshShortcut)
            shortCut.TargetPath = p_targetPath
            shortCut.WindowStyle = 1
            shortCut.Description = p_shortcutName
            shortCut.WorkingDirectory = p_workingDirectory
            shortCut.IconLocation = p_iconFile & ", " & p_iconNumber
            shortCut.Save()

            result = True
        Catch ex As System.Exception

        End Try

        Return result
    End Function
End Module
