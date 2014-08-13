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

Public Class Form1
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txt_watchpath = New System.Windows.Forms.TextBox
        Me.txt_folderactivity = New System.Windows.Forms.TextBox
        Me.btn_startwatch = New System.Windows.Forms.Button
        Me.btn_stop = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txt_watchpath
        '
        Me.txt_watchpath.Location = New System.Drawing.Point(40, 24)
        Me.txt_watchpath.Name = "txt_watchpath"
        Me.txt_watchpath.Size = New System.Drawing.Size(216, 20)
        Me.txt_watchpath.TabIndex = 0
        Me.txt_watchpath.Text = "TextBox1"
        '
        'txt_folderactivity
        '
        Me.txt_folderactivity.Location = New System.Drawing.Point(40, 80)
        Me.txt_folderactivity.Multiline = True
        Me.txt_folderactivity.Name = "txt_folderactivity"
        Me.txt_folderactivity.Size = New System.Drawing.Size(216, 112)
        Me.txt_folderactivity.TabIndex = 1
        Me.txt_folderactivity.Text = "TextBox1"
        '
        'btn_startwatch
        '
        Me.btn_startwatch.Location = New System.Drawing.Point(56, 216)
        Me.btn_startwatch.Name = "btn_startwatch"
        Me.btn_startwatch.TabIndex = 2
        Me.btn_startwatch.Text = "Button1"
        '
        'btn_stop
        '
        Me.btn_stop.Location = New System.Drawing.Point(160, 216)
        Me.btn_stop.Name = "btn_stop"
        Me.btn_stop.TabIndex = 3
        Me.btn_stop.Text = "Button2"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btn_stop)
        Me.Controls.Add(Me.btn_startwatch)
        Me.Controls.Add(Me.txt_folderactivity)
        Me.Controls.Add(Me.txt_watchpath)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btn_startwatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_startwatch.Click
        watchfolder = New System.IO.FileSystemWatcher

        'this is the path we want to monitor
        watchfolder.Path = txt_watchpath.Text

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
        btn_stop.Enabled = True

        'End of code for btn_start_click</PRE>
    End Sub

    Private Sub logchange(ByVal source As Object, ByVal e As _
                        System.IO.FileSystemEventArgs)
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
    End Sub

    Public Sub logrename(ByVal source As Object, ByVal e As _
                            System.IO.RenamedEventArgs)
   txt_folderactivity.Text &= "File" & e.OldName & _
                 " has been renamed to " & e.Name & vbCrLf
End Sub

    Private Sub btn_stop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_stop.Click
        ' Stop watching the folder
        watchfolder.EnableRaisingEvents = False
        btn_startwatch.Enabled = True
        btn_stop.Enabled = False
    End Sub
End Class
