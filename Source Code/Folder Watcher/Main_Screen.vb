'More Information 
'Use FileSystemWatcher.Filter Property to determine what files should be monitored, 
'eg setting filter property to "*.txt" shall monitor all the files with extension txt, 
'the default is *.* which means all the files with extension, if you want to monitor all 
'the files with and without extension please set the Filter property to "". 
'FileSystemWatcher can be used watch files on a local computer, a network drive, or a 
'remote computer but it does not raise events for CD. It only works on Windows 2000 and 
'Windows NT 4.0, common file system operations might raise more than one event. 
'For example, when a file is edited or moved, more than one event might be raised. 
'Likewise, some anti virus or other monitoring applications can cause additional events.

'The FileSystemWatcher will not watch the specified folder until the path property has 
'been set and EnableRaisingEvents is set to true. Set the FileSystemWatcher.IncludeSubdirectories 
'Property to true if you want to monitor subdirectories; otherwise, false. The default is 
'false. 

'FileSystemWatcher.Path property supports Universal Naming Convention (UNC) paths. If the 
'folder, which the path points is renamed the FileSystemWatcher reattaches itself to the 
'new renamed folder.


Imports System.IO
Imports System.Diagnostics

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Public watchfolder As FileSystemWatcher

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txt_watchpath As System.Windows.Forms.TextBox
    Friend WithEvents txt_folderactivity As System.Windows.Forms.TextBox
    Friend WithEvents btn_startwatch As System.Windows.Forms.Button
    Friend WithEvents btn_stop As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.txt_watchpath = New System.Windows.Forms.TextBox
        Me.txt_folderactivity = New System.Windows.Forms.TextBox
        Me.btn_startwatch = New System.Windows.Forms.Button
        Me.btn_stop = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txt_watchpath
        '
        Me.txt_watchpath.BackColor = System.Drawing.Color.LemonChiffon
        Me.txt_watchpath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt_watchpath.ForeColor = System.Drawing.Color.Black
        Me.txt_watchpath.Location = New System.Drawing.Point(24, 60)
        Me.txt_watchpath.Name = "txt_watchpath"
        Me.txt_watchpath.ReadOnly = True
        Me.txt_watchpath.Size = New System.Drawing.Size(352, 20)
        Me.txt_watchpath.TabIndex = 0
        Me.txt_watchpath.Text = ""
        '
        'txt_folderactivity
        '
        Me.txt_folderactivity.BackColor = System.Drawing.Color.LemonChiffon
        Me.txt_folderactivity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt_folderactivity.ForeColor = System.Drawing.Color.Black
        Me.txt_folderactivity.Location = New System.Drawing.Point(24, 88)
        Me.txt_folderactivity.Multiline = True
        Me.txt_folderactivity.Name = "txt_folderactivity"
        Me.txt_folderactivity.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt_folderactivity.Size = New System.Drawing.Size(352, 176)
        Me.txt_folderactivity.TabIndex = 1
        Me.txt_folderactivity.Text = ""
        '
        'btn_startwatch
        '
        Me.btn_startwatch.BackColor = System.Drawing.Color.Gold
        Me.btn_startwatch.Enabled = False
        Me.btn_startwatch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_startwatch.ForeColor = System.Drawing.Color.Black
        Me.btn_startwatch.Location = New System.Drawing.Point(24, 272)
        Me.btn_startwatch.Name = "btn_startwatch"
        Me.btn_startwatch.Size = New System.Drawing.Size(88, 23)
        Me.btn_startwatch.TabIndex = 2
        Me.btn_startwatch.Text = "Start Watching"
        '
        'btn_stop
        '
        Me.btn_stop.BackColor = System.Drawing.Color.Gold
        Me.btn_stop.Enabled = False
        Me.btn_stop.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_stop.ForeColor = System.Drawing.Color.Black
        Me.btn_stop.Location = New System.Drawing.Point(120, 272)
        Me.btn_stop.Name = "btn_stop"
        Me.btn_stop.Size = New System.Drawing.Size(88, 23)
        Me.btn_stop.TabIndex = 3
        Me.btn_stop.Text = "Stop Watching"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(400, 288)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Watcher Inactive"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select Folder to Monitor"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Yellow
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(384, 60)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 20)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Browse"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(456, 40)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "In order to monitor the file operations taking place is a specific folder, simply" & _
        " select a valid folder using the Browse button and then click on the Start Watch" & _
        "ing button to activate the Folder Watcher."
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(506, 312)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_stop)
        Me.Controls.Add(Me.btn_startwatch)
        Me.Controls.Add(Me.txt_folderactivity)
        Me.Controls.Add(Me.txt_watchpath)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(512, 344)
        Me.MinimumSize = New System.Drawing.Size(512, 344)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Folder Watcher"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btn_startwatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_startwatch.Click
        Try
            txt_folderactivity.Text &= "Watcher Activated at " & Format(Now(), "yyyy/MM/dd HH:mm:ss") & vbCrLf

            watchfolder = New System.IO.FileSystemWatcher

            'this is the path we want to monitor
            watchfolder.Path = txt_watchpath.Text
            watchfolder.IncludeSubdirectories = True


            'Add a list of Filter we want to specify
            'make sure you use OR for each Filter as we need to
            'all of those 

            watchfolder.NotifyFilter = IO.NotifyFilters.DirectoryName
            watchfolder.NotifyFilter = watchfolder.NotifyFilter Or _
                                       IO.NotifyFilters.FileName
            watchfolder.NotifyFilter = watchfolder.NotifyFilter Or _
                                       IO.NotifyFilters.Attributes

            ' add the handler to each event
            AddHandler watchfolder.Changed, AddressOf logchange
            AddHandler watchfolder.Created, AddressOf logchange
            AddHandler watchfolder.Deleted, AddressOf logchange

            ' add the rename handler as the signature is different
            AddHandler watchfolder.Renamed, AddressOf logrename

            'Set this property to true to start watching
            watchfolder.EnableRaisingEvents = True

            btn_startwatch.Enabled = False
            Button1.Enabled = False
            btn_stop.Enabled = True

            Label1.Text = "Watcher Active"
            Label1.ForeColor = Color.Green
            Label1.Refresh()

            'End of code for btn_start_click<
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub logchange(ByVal source As Object, ByVal e As _
                        System.IO.FileSystemEventArgs)
        Try
            If e.ChangeType = IO.WatcherChangeTypes.Changed Then
                txt_folderactivity.Text &= "File " & e.FullPath & _
                                        " has been modified" & vbCrLf
            End If
            If e.ChangeType = IO.WatcherChangeTypes.Created Then
                txt_folderactivity.Text &= "File " & e.FullPath & _
                                         " has been created" & vbCrLf
            End If
            If e.ChangeType = IO.WatcherChangeTypes.Deleted Then
                txt_folderactivity.Text &= "File " & e.FullPath & _
                                        " has been deleted" & vbCrLf
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub logrename(ByVal source As Object, ByVal e As _
                            System.IO.RenamedEventArgs)
        Try
            txt_folderactivity.Text &= "File" & e.OldName & _
                          " has been renamed to " & e.Name & vbCrLf
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub btn_stop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_stop.Click
        Try
            txt_folderactivity.Text &= "Watcher Deactivated at " & Format(Now(), "yyyy/MM/dd HH:mm:ss") & vbCrLf
            ' Stop watching the folder
            watchfolder.EnableRaisingEvents = False
            btn_startwatch.Enabled = True
            Button1.Enabled = True
            btn_stop.Enabled = False
            Label1.Text = "Watcher Inactive"
            Label1.ForeColor = Color.Red
            Label1.Refresh()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If txt_watchpath.Text.Length > 0 Then
                FolderBrowserDialog1.SelectedPath = txt_watchpath.Text
            End If
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
            If result = DialogResult.OK Then
                txt_watchpath.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub txt_watchpath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_watchpath.TextChanged
        Try
            If txt_watchpath.Text.Length > 0 Then
                btn_startwatch.Enabled = True
            Else
                btn_startwatch.Enabled = False
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Error_Handler(ByVal exc As Exception)
        MsgBox("The following error was trapped by Folder Watcher: " & vbCrLf & exc.ToString)
    End Sub
End Class
