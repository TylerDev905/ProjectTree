# ProjectTree
A component for VB .NET

<img src="http://i61.tinypic.com/30k4myt.png"/></img>


Project tree is a treeview that is extended with many features geared towards file management.

##Features
```
1. easy to implement
2. Live file updates
3. Rename files by editing the label
4. Custom coloured labels
5. Image toggling
6. Images display based on the file format
```

##Information
Using a fileworker we are able to watch the project directory/sub directories for updates. This will update the project tree live. By clicking on the item in the tree, the user can rename the file. Each file type can have a seperate image to be displayed. Toggling works by naming files in this format "folderOpen.png" / "folderClose.png". Methods are available for custom coloured labels. This is done by using a regex pattern and setting the brush colours for each capture group found. The rootpath must be set before the object can be initiated. The "Projects" directory should be created. This folder will be used to hold all project folders. The root path should also be the parent directory of the "Projects" Directory. A seperate context menu can be given to a project folder a directory folder or a file.

##Example
```vb
        Dim Projects as TreeViewExtended =  new TreeViewExtended()
        Projects.ImageList = Icons ' Imagelist for icons
        Projects.rootPath = RootPath ' Rootpath for the tree
        Projects.InitiateProjectTree(FilePath) ' institate
```
