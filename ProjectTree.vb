Imports System
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Contains many extra methods/properties for creating a treeView
''' </summary>
Public Class TreeViewExtended
    Inherits TreeView

    Public expandedNodes As Dictionary(Of String, Boolean)
    Public rootPath As String
    Public projectCmDir As ContextMenuStrip
    Public projectCmFile As ContextMenuStrip
    Public projectCmProject As ContextMenuStrip
    Public projectCount As Integer = 0
    Public projectFileCount As Integer = 0
    Public projectFolderCount As Integer = 0
    Public projectNodesCount As Integer = 0
    Public projectWatcher As FileSystemWatcher
    Public projectRootPath As String
    Public projectFocusedBgBrush As Brush
    Public projectFocusedCaptureBrushes() As Brush
    Public projectDefaultCaptureBrushes() As Brush

    ''' <summary>
    ''' Gets the nodes full filesystem path
    ''' </summary>
    ''' <param name="node">treeNode</param>
    Public ReadOnly Property nodeFullPath(node As TreeNode) As String
        Get
            Return Me.rootPath + "\Projects\" + node.FullPath
        End Get
    End Property

    ''' <summary>
    ''' Gets the nodes full filesystem path of the parent directory
    ''' </summary>
    ''' <param name="node">treeNode</param>
    Public ReadOnly Property nodeFullDirPath(node As TreeNode) As String
        Get
            If Path.GetExtension(node.FullPath) <> "" Then
                Dim name As String = "\" + Path.GetFileName(node.FullPath)
                Return (Me.rootPath + "\Projects\" + node.FullPath).Replace(name, "")
            Else
                Return Me.rootPath + "\Projects\" + node.FullPath
            End If

        End Get

    End Property

    ''' <summary>
    ''' Set some defaults for the new treeView
    ''' </summary>
    Sub New()
        projectDefaultCaptureBrushes = {Brushes.Black, Brushes.SkyBlue}
        projectFocusedBgBrush = Brushes.DodgerBlue
        projectFocusedCaptureBrushes = {Brushes.Black, Brushes.AliceBlue}
        Me.DoubleBuffered = True
    End Sub

    '############################################################################################
    '############################   TreeState   #################################################
    '############################################################################################

    ''' <summary>
    ''' Saves all expanded nodes in the tree to the expandedNodes dictionary
    ''' </summary>
    Public Sub SaveTreeExpandedNodes()
        expandedNodes = New Dictionary(Of String, Boolean)
        For Each node As TreeNode In Me.Nodes
            expandedNodes.Add(node.Name, node.IsExpanded)
            SaveExpandedNodes(node)
        Next
        projectNodesCount = expandedNodes.Count
    End Sub

    ''' <summary>
    ''' Saves expanded nodes to expandedNodes dictionary.
    ''' </summary>
    ''' <param name="nodes">Parent TreeNode</param>
    Private Sub SaveExpandedNodes(nodes As TreeNode)
        For Each node As TreeNode In nodes.Nodes
            expandedNodes.Add(node.Name, node.IsExpanded)
            SaveExpandedNodes(node)
        Next
    End Sub

    ''' <summary>
    ''' Load all expanded nodes in the tree to the expandedNodes dictionary
    ''' </summary>
    Public Sub LoadTreeExpandedNodes()
        For Each node As TreeNode In Me.Nodes
            If expandedNodes.Item(node.Name) Then
                node.Expand()
            Else
                node.Collapse()
            End If
            LoadExpandedNodes(node)
        Next
    End Sub

    ''' <summary>
    ''' Loads the saved expandedNodes dictionary.
    ''' </summary>
    ''' <param name="nodes">Parent TreeNode</param>
    Private Sub LoadExpandedNodes(nodes As TreeNode)
        For Each node As TreeNode In nodes.Nodes
            If expandedNodes.Item(node.Name) Then
                node.Expand()
            Else
                node.Collapse()
            End If
            LoadExpandedNodes(node)
        Next
    End Sub

    '############################################################################################
    '############################   ProjectTree   ###############################################
    '############################################################################################

    ''' <summary>
    ''' Populates the listTree with Project nodes from the directory given.
    ''' Each project node can contain directory and file nodes based on the filesystem.
    ''' [*] briefcaseClose is the image key and name for project nodes 
    ''' [*] folderClose is the image key and name for the directory nodes 
    ''' [remember] file image keys and names are generated using the file extension 
    ''' </summary>
    ''' <param name="newRootPath">The path that will be used to create the project nodes.</param>
    Public Sub InitiateProjectTree(newRootPath)
        With Me
            Me.DoubleBuffered = True
            rootPath = newRootPath
            projectRootPath = rootPath + "\Projects"

            'turn label editing on
            LabelEdit = True

            'create project nodes
            CreateProjectNodes()

            'Take over drawing for the label
            'DrawMode = TreeViewDrawMode.OwnerDrawText

            ' AddHandler DrawNode, AddressOf OnProjectDrawNode

            'Add some TreeViewExtended event handlers
            AddHandler .AfterLabelEdit, AddressOf OnProjectNodesLabelEdited
            AddHandler .BeforeLabelEdit, AddressOf Refresh
            AddHandler .AfterCollapse, AddressOf OnNodesToggle
            AddHandler .AfterExpand, AddressOf OnNodesToggle
        End With

        'create a file system watcher
        'this will sit on a seperate thread watching the projects directory and sub directories
        projectWatcher = New FileSystemWatcher()

        With projectWatcher

            .Path = Me.rootPath
            .IncludeSubdirectories = True

            'Sync to the gui so we do not have to delagate from a seperate thread.
            .SynchronizingObject = Me

            'Add handlers for watching Created,deleted and renamed events
            .NotifyFilter = (NotifyFilters.LastAccess Or NotifyFilters.LastWrite Or NotifyFilters.FileName Or NotifyFilters.DirectoryName Or NotifyFilters.Size Or NotifyFilters.CreationTime)
            AddHandler .Created, AddressOf OnProjectWatcherCreated
            AddHandler .Deleted, AddressOf OnProjectWatcherDeleted
            AddHandler .Renamed, AddressOf OnProjectWatcherRenamed

            'Start the file watcher
            .EnableRaisingEvents = True

        End With
    End Sub


    Public Sub loadCodeTreeFiles()

    End Sub

    ''' <summary>
    ''' Clears the tree, recreates project nodes and refreshes the drawing.
    ''' </summary>
    Public Sub RefreshProjectNodes()
        With Me
            .Nodes.Clear()
            CreateProjectNodes()
            Me.Refresh()
        End With
    End Sub

    ''' <summary>
    ''' Looks for directories in the project folder.
    ''' For each directory found a project node will be
    ''' created in the project tree
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateProjectNodes()
        With Me
            .Refresh()
            For Each project As String In My.Computer.FileSystem.GetDirectories(.projectRootPath)
                .Nodes.Add(project, Path.GetFileName(project), "briefcaseClose", "briefcaseClose")
                .Nodes(project).ContextMenuStrip = projectCmProject
                CreateFileAndFolderNodes(.Nodes(project))
                projectCount = projectCount + 1
            Next
        End With
    End Sub

    ''' <summary>
    ''' create directories and file nodes
    ''' recursize method
    ''' </summary>
    ''' <param name="node">parent tree node to attach child nodes</param>
    Public Sub CreateFileAndFolderNodes(node As TreeNode)
        With node
            'create all directory nodes
            For Each dirItem As String In My.Computer.FileSystem.GetDirectories(node.Name)
                .Nodes.Add(dirItem, dirItem.Replace(.Name, "").Replace("\", ""), "folderClose", "folderClose")
                .Nodes(dirItem).ContextMenuStrip = projectCmDir
                CreateFileAndFolderNodes(.Nodes(dirItem))
                projectFolderCount = projectFolderCount + 1
            Next
            'create all file nodes
            For Each fileItem As String In My.Computer.FileSystem.GetFiles(node.Name)
                Dim extension = Path.GetExtension(fileItem).Replace(".", "")
                .Nodes.Add(fileItem, Path.GetFileName(fileItem), extension, extension)
                .Nodes(fileItem).ContextMenuStrip = projectCmFile
                projectFileCount = projectFileCount + 1
            Next


        End With

    End Sub

    ''' <summary>
    ''' Renames a file
    ''' </summary>
    ''' <param name="originalPath">The path of the file to rename</param>
    ''' <param name="newName">the new name</param>
    Private Sub RenameFile(originalPath As String, newName As String)
        If My.Computer.FileSystem.FileExists(originalPath) Then
            Dim originalName = Path.GetFileName(originalPath)
            If My.Computer.FileSystem.FileExists(originalPath.Replace(originalName, newName)) = False Then
                My.Computer.FileSystem.RenameFile(originalPath, newName)
            Else
                MsgBox("An error occured. A file by that name already exists in this directory.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Renames a directory
    ''' </summary>
    ''' <param name="originalPath">The path to the directory you would like to rename.</param>
    ''' <param name="originalName">The original name of the file.</param>
    ''' <param name="newName">The new name of the file.</param>
    Private Sub RenameDirectory(originalPath As String, originalName As String, newName As String)
        Try
            If My.Computer.FileSystem.DirectoryExists(originalPath) Then
                If My.Computer.FileSystem.DirectoryExists(originalPath.Replace(originalName, newName)) = False Then
                    My.Computer.FileSystem.RenameDirectory(originalPath, newName)
                Else
                    MsgBox("An error occured. A directory by that name already exists in this directory.")
                End If
            End If
        Catch ex As Exception
            MsgBox("Error no access please ensure you are running in administration mode and antivirus software is setup correctly.")
        End Try
    End Sub

    '############################################################################################
    '############################   Label Drawing Helper Methods   ##############################
    '############################################################################################

    ''' <summary>
    ''' Gets the font used on the node if not defaults to the TreeViewExtended Font
    ''' </summary>
    ''' <param name="e">DrawTreeNodeEventArgs</param>
    ''' <returns>Font resource</returns>
    Public Function ColorLabelGetFont(e As DrawTreeNodeEventArgs)
        Dim labelFont As Font = e.Node.NodeFont
        If labelFont Is Nothing Then
            labelFont = Me.Font
        End If
        Return labelFont
    End Function

    ''' <summary>
    ''' Used to color a labels background
    ''' </summary>
    ''' <param name="labelBrush">Drawing brush</param>
    ''' <param name="e">DrawTreeNodeEventArgs</param>
    Public Sub ColorLabelBg(labelBrush As Brush, e As DrawTreeNodeEventArgs)
        Dim bounds As Rectangle = e.Bounds
        bounds.Width = bounds.Width - 2
        e.Graphics.FillRectangle(labelBrush, bounds)
    End Sub

    ''' <summary>
    ''' Using the capture groups from the pattern given to color a label 
    ''' Brush Color is assigned per capture group
    ''' </summary>
    ''' <param name="labelText"></param>
    ''' <param name="pattern"></param>
    ''' <param name="labelBrushes"></param>
    ''' <param name="e">DrawTreeNodeEventArgs</param>
    ''' <returns>The end of the labels position</returns>
    Public Function ColorLabelText(labelText As String, pattern As String, labelBrushes() As Brush, e As DrawTreeNodeEventArgs)
        Dim labelFont As Font = ColorLabelGetFont(e) 'gets the font
        Dim leftPosition As Integer = e.Bounds.Left 'gets the bounds of the node

        'matches the pattern against the labels text to get the capture groups
        Dim captured As Match = Regex.Match(labelText, pattern, RegexOptions.IgnoreCase)
        Dim index As Integer = 0
        Dim firstRun As Boolean = False

        'for every capture besides first run
        'draw the string with the brush color chosen
        For Each capture As Group In captured.Groups
            If firstRun Then
                e.DrawDefault = False
                e.Graphics.DrawString(capture.Value, labelFont, labelBrushes(index), leftPosition, e.Bounds.Top)
                leftPosition = (leftPosition + e.Graphics.MeasureString(capture.Value, labelFont).Width) - 5
                index = index + 1
            Else
                firstRun = True
            End If
        Next
        Return firstRun
    End Function


    '############################################################################################
    '############################   Event Handlers   ############################################
    '############################################################################################


    ''' <summary>
    ''' Checks the images name for toggle instructions.  
    ''' </summary>
    ''' <param name="sender">TreeViewExtended</param>
    ''' <param name="e">TreeViewEventArgs</param>
    ''' <remarks>Expanded:imagenameOpen.png Collapsed:imagenameClose.png</remarks>
    Public Sub OnNodesToggle(sender As Object, e As TreeViewEventArgs)
        With e
            Me.Refresh()
            If .Node.ImageKey.Contains("Open") Then
                'The Open command was found in the file name
                .Node.ImageKey = .Node.ImageKey.Replace("Open", "Close")
                ''change the file name with the command Close
                .Node.SelectedImageKey = .Node.ImageKey.Replace("Open", "Close")

            ElseIf .Node.ImageKey.Contains("Close") Then
                'the Close command was found in the file name
                .Node.ImageKey = .Node.ImageKey.Replace("Close", "Open")
                'change the file name with the command Open
                .Node.SelectedImageKey = .Node.ImageKey.Replace("Close", "Open")
            End If
        End With
    End Sub

    ''' <summary>
    ''' When the label is changed we rename the item in the file system
    ''' </summary>
    ''' <param name="sender">TreeViewExtended</param>
    ''' <param name="e">NodeLabelEditEventArgs</param>
    Private Sub OnProjectNodesLabelEdited(sender As Object, e As NodeLabelEditEventArgs)
        If e.Label <> Nothing Then
            RenameDirectory(e.Node.Name, e.Node.Text, e.Label)
            RenameFile(e.Node.Name, e.Label)
        End If
    End Sub

    ''' <summary>
    ''' When a new file is created or found in the file system.
    ''' Update the project tree with the new node.
    ''' </summary>
    ''' <param name="sender">TreeViewExtended</param>
    ''' <param name="e">FileSystemEventArgs</param>
    Private Sub OnProjectWatcherCreated(sender As Object, e As FileSystemEventArgs)
        Try
            With Me
                Dim fullPath = e.FullPath.Replace("\" + Path.GetFileName(e.FullPath), "")
                Dim nodes As Array = .Nodes.Find(fullPath, True)
                Dim ext = Path.GetExtension(e.FullPath).Replace(".", "")

                If ext = "" Then
                    Me.Nodes.Add(e.FullPath, Path.GetFileName(e.FullPath), "briefcaseClose", "briefcaseClose")
                    Me.Nodes(e.FullPath).ContextMenuStrip = .projectCmDir
                Else
                    Dim item As TreeNode = nodes(0)
                    item.Nodes.Add(e.FullPath, Path.GetFileName(e.FullPath), ext, ext)
                    item.Nodes(e.FullPath).ContextMenuStrip = .projectCmFile
                End If

            End With
        Catch ex As Exception
            Console.WriteLine("An exception was thrown when trying to create the selected item.")
        End Try
    End Sub

    ''' <summary>
    ''' A file was deleted from the file structure.
    ''' Delete the file node from the project tree.
    ''' </summary>
    ''' <param name="sender">TreeViewExtended</param>
    ''' <param name="e">FileSystemEventArgs</param>
    Private Sub OnProjectWatcherDeleted(sender As Object, e As FileSystemEventArgs)
        Try
            With Me
                Dim fullPath = e.FullPath
                Dim nodes As Array = .Nodes.Find(fullPath, True)
                Dim node As TreeNode = nodes(0)
                node.Remove()
            End With
        Catch ex As Exception
            Console.WriteLine("An exception was thrown when trying to delete the selected item.")
        End Try
    End Sub

    ''' <summary>
    ''' When a project, folder or file is renamed in the filesystem.
    ''' Update the node to the new name.  
    ''' </summary>
    ''' <param name="sender">TreeViewExtended</param>
    ''' <param name="e">RenamedEventArgs</param>
    Private Sub OnProjectWatcherRenamed(sender As Object, e As RenamedEventArgs)
        Try
            With Me
                If My.Computer.FileSystem.FileExists(e.FullPath) Then
                    Dim nodes As Array = .Nodes.Find(e.OldFullPath, True)
                    Dim node As TreeNode = nodes(0)
                    node.Name = e.FullPath
                    node.Text = Path.GetFileName(e.FullPath)
                Else
                    Dim nodes As Array = .Nodes.Find(e.OldFullPath, True)
                    Dim node As TreeNode = nodes(0)
                    node.Name = e.FullPath
                    node.Text = Path.GetFileName(e.FullPath)
                End If
            End With
        Catch ex As Exception
            Console.WriteLine("An exception was thrown when trying to rename the selected item.")
        End Try
    End Sub



End Class



