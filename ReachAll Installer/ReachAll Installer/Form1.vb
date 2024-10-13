Imports System
Imports System.IO

Public Class Form1
    Const SLEEP_TIME_MS As Integer = 250
    Const VERSION As Integer = 3
    Dim protectedFiles() As String = {"config.xml", "callrecordingadmindata.xml", "web.config"}

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub StartInstallationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StartInstallationToolStripMenuItem1.Click
        ' Install the core ReachAll application ?
        Dim looping As Boolean = True
        Dim installing As Boolean = True

        Label1.Text = ""

        While looping
            looping = False

            Select Case MsgBox("Do you want to install the core ReachAll application ?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                    InstallCoreReachAllApplication()

                Case MsgBoxResult.No
                    ' Goto next component

                Case MsgBoxResult.Cancel
                    Select Case MsgBox("Are you sure you want to stop the installation process ?", MsgBoxStyle.YesNo)
                        Case MsgBoxResult.Yes
                            installing = False

                        Case MsgBoxResult.No
                            looping = True
                    End Select
            End Select
        End While

        If installing Then
            ' Install the ReachAll web interface ?
            looping = True

            While looping
                looping = False

                Select Case MsgBox("Do you want to install the ReachAll Web Interface ?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        Dim myPath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles).TrimEnd("\") & "\ReachAll"
                        Dim myUrl As String = "http://" & Environment.MachineName & "/ReachAllWebInterface/ReachAllWebInterface.aspx"

                        GenericInstall("c:\inetpub\wwwroot", "ReachAllWebInterface", "ReachAll Web Interface")
                        SetXMLField("c:\inetpub\wwwroot\ReachAllWebInterface\config.xml", "reachAllApplicationPath", myPath)
                        SetXMLField("c:\inetpub\wwwroot\ReachAllWebInterface\config.xml", "ourHomePage", myUrl)
                        CreateFavourite("ReachAll Web Interface", myUrl)

                        If VERSION = 3 Then
                            ' Remove the style sheets from the application folder as the application will use the web interface versions
                            Dim myFileNamePath As String = GetTargetApplicationPath().GetPath

                            If myFileNamePath <> "" Then
                                Dim myFileName As String = myFileNamePath & "\ReachAllReportStyleSheet.css"

                                If FileExists(myFileName) Then My.Computer.FileSystem.DeleteFile(myFileName)

                                myFileName = myFileNamePath & "\ReachAllReportStyleSheetNewInterface.css"

                                If FileExists(myFileName) Then My.Computer.FileSystem.DeleteFile(myFileName)
                            End If
                        End If

                    Case MsgBoxResult.No
                        ' Goto next component

                    Case MsgBoxResult.Cancel
                        Select Case MsgBox("Are you sure you want to stop the installation process ?", MsgBoxStyle.YesNo)
                            Case MsgBoxResult.Yes
                                installing = False

                            Case MsgBoxResult.No
                                looping = True
                        End Select
                End Select
            End While
        End If

        If installing Then
            ' Install the call recording interface ?
            looping = True

            While looping
                looping = False

                Select Case MsgBox("Do you want to install the ReachAll Call Recording Interface ?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        Dim myUrl As String = GetCallRecordingURL()
                        Dim myPath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles).TrimEnd("\") & "\ReachAll"
                        Dim myOptionsFilename As String = myPath & "\ReachAllOptionSettings.txt"

                        GenericInstall("c:\inetpub\wwwroot", "CallRecordingGUI", "ReachAll Call Recording Interface")
                        CreateFavourite("ReachAll Call Recording Interface", myUrl)
                        SetOptionField(myOptionsFilename, GetCallRecordingURLOptionHeading(), myUrl)

                    Case MsgBoxResult.No
                        ' Goto next component

                    Case MsgBoxResult.Cancel
                        Select Case MsgBox("Are you sure you want to stop the installation process ?", MsgBoxStyle.YesNo)
                            Case MsgBoxResult.Yes
                                installing = False

                            Case MsgBoxResult.No
                                looping = True
                        End Select
                End Select
            End While
        End If

        Label1.Text = "Installation complete !"
        Label2.Visible = False
        ProgressBar1.Visible = False
    End Sub

    Private Function GetTargetApplicationPath() As WindowsPathClass
        Dim myPath As New WindowsPathClass("C:\Program Files (x86)")
        Dim x As New IO.DirectoryInfo(myPath.GetPath)
        Dim parentPathExists As Boolean = False
        Dim myPathAsString = ""

        myPath.SetPath(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
        x = Nothing
        x = New IO.DirectoryInfo(myPath.GetPath)

        If x.Exists Then parentPathExists = True

        If parentPathExists Then
            myPath.AddPath("ReachAll")
        Else
            myPath = Nothing
        End If

        Return myPath
    End Function

    Private Sub InstallCoreReachAllApplication()
        ' Establish the parent directory for the application directory
        Dim myPath As New WindowsPathClass("C:\Program Files (x86)")
        Dim x As New IO.DirectoryInfo(myPath.GetPath)
        Dim parentPathExists As Boolean = False

        '        If x.Exists Then
        'parentPathExists = True
        '     Else
        '    myPath = "C:\Program Files"
        '    x = Nothing
        '  x = New IO.DirectoryInfo(myPath)

        '   If x.Exists Then parentPathExists = True
        '   End If

        myPath.SetPath(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
        x = Nothing
        x = New IO.DirectoryInfo(myPath.GetPath)

        If x.Exists Then parentPathExists = True

        If parentPathExists Then
            Dim errorOccured As Boolean = False

            ' Does application directory already exist ?
            myPath.AddPath("ReachAll")
            x = Nothing
            x = New IO.DirectoryInfo(myPath.GetPath)

            If x.Exists Then
                ' Application directory already exists
                MsgBox("Target folder " & myPath.GetPath & " already exists")
            Else
                ' Create the application directory
                Try
                    IO.Directory.CreateDirectory(myPath.GetPath)
                Catch e As Exception
                    MsgBox("Could not create folder " & myPath.GetPath)
                    errorOccured = True
                End Try

                If Not errorOccured Then
                    MsgBox("Folder " & myPath.GetPath & " was created")
                End If
            End If

            If Not errorOccured Then
                '  If Not myPath.EndsWith("\") Then myPath &= "\"

                ' Copy the files over
                For Each filename As String In reachallApplicationFileList
                    Dim destFilename As String = filename

                    ' Remove any path information to leave only the actual filename
                    If destFilename.Contains("\") Then destFilename = destFilename.Substring(destFilename.LastIndexOf("\") + 1)

                    Label1.Text = "Copying " & filename & " to " & myPath.CreateFullName(destFilename)
                    Me.Refresh()

                    Try
                        FileCopy(GetApplicationPath().CreateFullName(filename), myPath.CreateFullName(destFilename))
                    Catch e As Exception
                        MsgBox("File copy error: " & e.Message)
                    End Try

                    Sleep()
                Next

                Dim shortCutParent As String = GetShortCutParent()

                If IO.Directory.Exists(shortCutParent) Then
                    Label1.Text = "Creating shortcut ..."
                    CreateShortCut("ReachAll Viewer", shortCutParent & "\ReachAll", myPath.CreateFullName("ReachAll Viewer.exe"), myPath.GetPath, myPath.CreateFullName("iconalpha.ico"), 0)
                    Sleep()
                    Label1.Text = ""
                Else
                    MsgBox("Cannot locate shortcut directory " & shortCutParent)
                End If
            End If

            ' Copy the SQL folder and files over
            If Not errorOccured Then
                x = Nothing
                x = New IO.DirectoryInfo(myPath.GetPath & "\SQL")

                If x.Exists Then
                    ' SQL directory already exists
                    MsgBox("SQL folder " & myPath.GetPath & "\SQL already exists")
                Else
                    ' Create the application directory
                    Try
                        IO.Directory.CreateDirectory(myPath.GetPath & "\SQL")
                    Catch e As Exception
                        MsgBox("Could not create folder " & myPath.GetPath & "\SQL")
                        errorOccured = True
                    End Try

                    If Not errorOccured Then
                        MsgBox("Folder " & myPath.GetPath & "\SQL was created")
                    End If
                End If

                Sleep()

                If Not errorOccured Then
                    Dim myFiles = Directory.GetFiles("SQL")

                    For Each filename As String In myFiles
                        Label1.Text = "Copying " & filename & " to " & myPath.CreateFullName(filename)
                        Me.Refresh()

                        Try
                            FileCopy(GetApplicationPath().CreateFullName(filename), myPath.CreateFullName(filename))
                        Catch e As Exception
                            MsgBox("File copy error: " & e.Message)
                        End Try

                        Sleep()
                    Next
                End If
            End If
        Else
            ' Cannot install as parent path does not exist
            MsgBox("Error: " & myPath.GetPath & " does not exist - install of Core ReachAll Application has been stopped", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub DeleteShortCut()
        Dim shortCutParent As String = GetShortCutParent()

        If IO.Directory.Exists(shortCutParent) Then
            Dim myFileName As String = shortCutParent & "\ReachAll\ReachAll Viewer.lnk"

            If FileExists(myFileName) Then My.Computer.FileSystem.DeleteFile(myFileName)
        End If
    End Sub

    Private Function GetShortCutParent() As String
        Dim shortCutParent As String = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"

        If Not IO.Directory.Exists(shortCutParent) Then shortCutParent = "C:\Documents And Settings\All Users\Start Menu\Programs"

        Return shortCutParent
    End Function

    Private Sub CopyFile(ByRef p_source As String, ByRef p_dest As String)
        FileCopy(p_source, p_dest)
    End Sub

    Private Sub GenericInstall(ByRef p_root As String, ByRef p_applicationDirectory As String, ByRef p_title As String)
        ' Establish the parent directory for the application directory
        Dim myPath As String = p_root
        Dim x As New IO.DirectoryInfo(myPath)

        ' Count the files to copy
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = DeepCopy(GetApplicationPath().CreateFullName(p_applicationDirectory), "")

        If x.Exists Then
            Dim errorOccured As Boolean = False

            ' Does application directory already exist ?
            myPath &= "\" & p_applicationDirectory
            x = Nothing
            x = New IO.DirectoryInfo(myPath)

            If x.Exists Then
                ' Application directory already exists
                MsgBox("Target folder " & myPath & " already exists")
            Else
                ' Create the application directory
                Try
                    IO.Directory.CreateDirectory(myPath)
                Catch e As Exception
                    MsgBox("Could not create folder " & myPath)
                    errorOccured = True
                End Try

                If Not errorOccured Then
                    MsgBox("Folder " & myPath & " was created")
                End If
            End If

            If Not errorOccured Then
                ' Copy the files over
                ProgressBar1.Value = 0
                DeepCopy(GetApplicationPath().CreateFullName(p_applicationDirectory), myPath)
            End If
        Else
            ' Cannot install as parent path does not exist
            MsgBox("Error: " & myPath & " does not exist - install of " & p_title & " has been stopped", MsgBoxStyle.Critical)
        End If

        ProgressBar1.Value = 0
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.Text = ""
        Initialise()
    End Sub

    Private Function DeepCopy(ByRef p_source As String, ByRef p_dest As String) As Integer
        Dim numberOfFiles As Integer = 0
        Dim myBaseDirectoryInfo As New IO.DirectoryInfo(p_source)
        Dim myFileInfoArray As IO.FileInfo() = myBaseDirectoryInfo.GetFiles()
        Dim myDirectoryInfoArray As IO.DirectoryInfo() = myBaseDirectoryInfo.GetDirectories

        For Each myFileInfo As IO.FileInfo In myFileInfoArray
            ' Ignore logfiles
            If Not myFileInfo.Name.ToLower.EndsWith("logfile.txt") Then
                ' Are we copying, or just counting ?

                If p_dest.Length > 0 Then
                    ' If the config file already exists, check if the user wants to overwrite
                    Dim myTarget As String = p_dest & "\" & myFileInfo.Name
                    Dim x As New IO.FileInfo(myTarget)
                    Dim writeIt As Boolean = True

                    If protectedFiles.Contains(myFileInfo.Name.ToLower) Then
                        If x.Exists Then
                            If MsgBox("Do you want to overwrite " & x.Name & " ?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then writeIt = False
                        End If
                    End If

                    If writeIt Then
                        Label1.Text = "Copying " & myFileInfo.Name & " to " & myTarget
                        Me.Refresh()
                        FileCopy(myFileInfo.FullName, myTarget)
                        Sleep()
                    End If
                End If

                numberOfFiles += 1

                If p_dest.Length > 0 Then ProgressBar1.Value += 1
            End If
        Next

        Label1.Text = ""

        For Each myDirectoryInfo As IO.DirectoryInfo In myDirectoryInfoArray
            If p_dest.Length > 0 Then
                ' Do we need to create the dest directory ?
                Dim errorOccured As Boolean = False
                Dim x As New IO.DirectoryInfo(p_dest & "\" & myDirectoryInfo.Name)

                If Not x.Exists Then
                    Try
                        IO.Directory.CreateDirectory(p_dest & "\" & myDirectoryInfo.Name)
                    Catch ex As Exception
                        MsgBox("Cannot create folder " & p_dest & "\" & myDirectoryInfo.Name)
                        errorOccured = True
                    End Try
                End If

                If Not errorOccured Then DeepCopy(myDirectoryInfo.FullName, p_dest & "\" & myDirectoryInfo.Name)
            Else
                numberOfFiles += DeepCopy(myDirectoryInfo.FullName, "")
            End If
        Next

        Return numberOfFiles
    End Function

    Private Sub Sleep(Optional ByVal p_delayMs As Integer = SLEEP_TIME_MS)
        System.Threading.Thread.Sleep(p_delayMs)
    End Sub

    Private Sub SetXMLField(ByRef p_filename As String, ByRef p_fieldName As String, ByRef p_value As String)
        ' Does file exist ?
        If IO.File.Exists(p_filename) Then
            ' Yes. Open for reading
            Dim myReader As New IO.StreamReader(p_filename)
            Dim reading As Boolean = True
            Dim myMatchString = "<" & p_fieldName.ToLower & ">"
            Dim myMatchEndString = "</" & p_fieldName.ToLower & ">"
            Dim myLines As New List(Of String)

            While reading
                Dim myLine As String = myReader.ReadLine

                If myLine Is Nothing Then
                    reading = False
                Else
                    ' Does this line contain the field ?
                    If myLine.ToLower.Contains(myMatchString) And myLine.ToLower.Contains(myMatchEndString) Then
                        Dim myNewLine As String = myLine.Substring(0, myLine.ToLower.IndexOf(myMatchString))

                        myNewLine &= "<" & p_fieldName & ">" & p_value
                        myNewLine &= myLine.Substring(myLine.ToLower.IndexOf(myMatchEndString))
                        myLines.Add(myNewLine)
                    Else
                        myLines.Add(myLine)
                    End If
                End If
            End While

            myReader.Close()
            myReader = Nothing

            Dim myWriter As New IO.StreamWriter(p_filename)

            For i = 0 To myLines.Count - 1
                myWriter.WriteLine(myLines(i))
            Next

            myWriter.Close()
            myWriter = Nothing
        End If
    End Sub

    Private Sub CreateFavourite(ByRef p_title As String, ByRef p_url As String)
        Dim favouritesFolder As String = System.Environment.GetFolderPath(Environment.SpecialFolder.Favorites)
        Dim myWriter As New IO.StreamWriter(favouritesFolder.TrimEnd("\") & "\" & p_title & ".url")

        myWriter.WriteLine("[InternetShortcut]")
        myWriter.WriteLine("URL=" & p_url)
        myWriter.Close()
        myWriter = Nothing
    End Sub

    Private Sub RemoveFavourite(ByRef p_title As String)
        Dim favouritesFolder As String = System.Environment.GetFolderPath(Environment.SpecialFolder.Favorites)
        Dim myFileName As String = favouritesFolder.TrimEnd("\") & "\" & p_title & ".url"

        Try
            My.Computer.FileSystem.DeleteFile(myFileName)
        Catch ex As Exception
            MsgBox("Could not delete shortcut filename: " & myFileName)
        End Try
    End Sub

    Private Sub UninstallApplicationAndService()
        ' Ask the user to stop and remove the ReachAll service
        If MsgBox("Please stop and remove the ReachAll service", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            ' Delete the application folder
            Dim myPath As String = GetTargetApplicationPath().GetPath

            DeleteShortCut()

            If IO.Directory.Exists(myPath) Then My.Computer.FileSystem.DeleteDirectory(myPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Else
            MsgBox("ReachAll uninstallation terminated ..")
        End If
    End Sub

    Private Sub UninstallReportingWebInterface()
        Dim myDirectory As String = "C:\inetpub\wwwroot\ReachAllWebInterface"

        RemoveFavourite("ReachAll Web Interface")

        If Directory.Exists(myDirectory) Then My.Computer.FileSystem.DeleteDirectory(myDirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub

    Private Sub UninstallCallRecordingInterface()
        Dim myDirectory As String = "C:\inetpub\wwwroot\CallRecordingGUI"

        RemoveFavourite("ReachAll Call Recording Interface")

        If Directory.Exists(myDirectory) Then My.Computer.FileSystem.DeleteDirectory(myDirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub

    Private Sub UninstallToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UninstallToolStripMenuItem.Click
        Label1.Text = ""

        If MsgBox("Do you want to remove the ReachAll application and service ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            UninstallApplicationAndService()
        End If

        If MsgBox("Do you want to remove the ReachAll Reporting Web Interface ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            UninstallReportingWebInterface()
        End If

        If MsgBox("Do you want to remove the ReachAll Call Recording Interface ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            UninstallCallRecordingInterface()
        End If

        MsgBox("Uninstallation complete ..")
    End Sub

    Private Sub To30ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles To30ToolStripMenuItem.Click
        Dim running As Boolean = True

        Label1.Text = ""

        If MsgBox("Do you want to upgrade the ReachAll application and service to 3.0 ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim myPath As WindowsPathClass = GetTargetApplicationPath()
            Dim localFileList As New List(Of String)

            MsgBox("Please stop and remove the ReachAll service")

            My.Computer.FileSystem.DeleteFile(myPath.GetPath & "\reachAllReportStyleSheet.css")

            ' Copy the files over
            localFileList.Add("ReachAll Viewer.exe")
            localFileList.Add("ReachAllService.exe")
            localFileList.Add("ReachAllWebInterface\ReachAllReportStyleSheet.css")
            localFileList.Add("ReachAllWebInterface\ReachAllReportStyleSheetNewInterface.css")

            For Each filename As String In localFileList
                Dim destFilename As String = filename

                ' Remove any path information to leave only the actual filename
                If destFilename.Contains("\") Then destFilename = destFilename.Substring(destFilename.LastIndexOf("\") + 1)

                Label1.Text = "Copying " & filename & " to " & myPath.CreateFullName(destFilename)
                Me.Refresh()

                Try
                    FileCopy(GetApplicationPath().CreateFullName(filename), myPath.CreateFullName(destFilename))
                Catch ex As Exception
                    MsgBox("File copy error: " & ex.Message)
                End Try

                Sleep()
            Next

            MsgBox("Please reinstall and start the ReachAll service")
        End If

        If MsgBox("Do you want to upgrade the ReachAll Web Interface to 3.0 ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim filesToDelete As New List(Of String)
            Dim localFileList As New List(Of String)
            Dim cssFolderDeleted As Boolean = True

            filesToDelete.Add("bootstrap.min.css")
            filesToDelete.Add("bootstrapStyleSheet.css")
            filesToDelete.Add("reachAllStyleSheet.css")
            filesToDelete.Add("ReachAllWebInterfaceStyleSheet.css")
            filesToDelete.Add("defaultStyleSheet.css")

            For Each filename As String In filesToDelete
                Dim fileDeleted As Boolean = True

                Label1.Text = "Deleting " & filename
                Me.Refresh()

                Try
                    My.Computer.FileSystem.DeleteFile("C:\inetpub\wwwroot\ReachAllWebInterface\" & filename)
                Catch ex As Exception
                    fileDeleted = False

                    If MsgBox("File delete error: " & ex.Message, MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                        running = False
                        Exit For
                    End If
                End Try

                If running And fileDeleted Then Sleep()
            Next

            Try
                My.Computer.FileSystem.DeleteDirectory("C:\inetpub\wwwroot\ReachAllWebInterface\css", FileIO.DeleteDirectoryOption.DeleteAllContents)
            Catch ex As Exception
                cssFolderDeleted = False

                If MsgBox("Folder delete error: " & ex.Message, MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                    running = False
                End If
            End Try

            If cssFolderDeleted Then
                Try
                    My.Computer.FileSystem.CreateDirectory("C:\inetpub\wwwroot\ReachAllWebInterface\css")
                Catch ex As Exception
                    cssFolderDeleted = False

                    If MsgBox("Folder delete error: " & ex.Message, MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                        running = False
                    End If
                End Try
            End If

            If running Then
                Dim myPath As WindowsPathClass = GetTargetApplicationPath()
                Dim filesToCheck As New List(Of String)

                GenericInstall("c:\inetpub\wwwroot", "ReachAllWebInterface", "ReachAll Web Interface")

                filesToCheck.Add("ReachAllReportStyleSheet.css")
                filesToCheck.Add("ReachAllReportStyleSheetNewInterface.css")

                For Each fileToCheck As String In filesToCheck
                    Dim lookingFor As String = "C:\inetpub\wwwroot\ReachAllWebInterface\" & fileToCheck

                    If FileExists(lookingFor) Then
                        Dim removing As String = myPath.CreateFullName(fileToCheck)

                        Label1.Text = "Removing " & removing & " as it already exists in C:\inetpub\wwwroot\ReachAllWebInterface"

                        Try
                            My.Computer.FileSystem.DeleteFile(removing)
                        Catch ex As Exception
                            MsgBox("File delete error: " & ex.Message)
                        End Try

                        Sleep()
                    End If
                Next
            End If

        End If

        If running Then
            If MsgBox("Do you want to upgrade the ReachAll Call Recording Interface to 3.0 ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                GenericInstall("c:\inetpub\wwwroot", "CallRecordingGUI", "ReachAll Call Recording Interface")
            End If

            Label1.Text = "Upgrade complete .."
        Else
            Label1.Text = "Upgrade cancelled .."
        End If

        Me.Refresh()
    End Sub
End Class
