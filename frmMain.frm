VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Favorites"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Open project with Visual Basic"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Remove from favorites"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton NewFolder 
      Caption         =   "New Folder"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "Create new folder node"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Exit and save favorites"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Exit without saving favorites"
      Top             =   2760
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imglstSamll 
      Left            =   3240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "Favorite"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E64
            Key             =   "FldrClose"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FBE
            Key             =   "FldrOpen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1118
            Key             =   "Project"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFavorites 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      ImageList       =   "imglstSamll"
      Appearance      =   1
   End
   Begin VB.Label lblCap 
      Caption         =   "Drag and drop project files."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ToDo Sort tree

Dim cProject As clsProperties
Dim m_nodX As Node
Dim m_sPath As String
Dim m_bDragNode As Boolean
Dim m_bSaveFavorites As Boolean

Private Sub cmdCancel_Click()

    ' Don't save.
    m_bSaveFavorites = False
    
    Unload Me
    
End Sub

Private Sub cmdLaunch_Click()

    Dim sPath As String
        
    If IsProjectNode(tvFavorites.SelectedItem) Then
        sPath = tvFavorites.SelectedItem.Tag
        RunAssociated sPath
        
    Else
        MsgBox "Select a project to launch"
        
    End If
    
End Sub

Private Sub cmdRemove_Click()

    'remove selected node.
    
    Dim nodRemove As Node
    
    ' reference the node to remove
    Set nodRemove = tvFavorites.SelectedItem
    
    ' is node root?
    If nodRemove.Key <> nodRemove.Root.Key Then
        ' get confirmation.
        If MsgBox("Remove Node?", vbExclamation & vbOKCancel) = vbOK Then
            tvFavorites.Nodes.Remove nodRemove.Index
            m_bSaveFavorites = True
            
        End If
        
    End If
    
    tvFavorites.SetFocus
    
End Sub

Private Sub Form_Initialize()

    ' establise path to favorites text file.
    
    Const sFile As String = "VBFavorites.txt"

    'parse path to paths.txt
    If Len(App.Path) > 4 Then  ' Apps in folder add backslash.
        m_sPath = App.Path & "\" & sFile

    Else  ' Apps in the root no backslash.
        m_sPath = App.Path & sFile

    End If

    With tvFavorites
        .OLEDropMode = ccOLEDropManual
        .HideSelection = False
    End With
    
End Sub

Private Sub Form_Load()

    Dim nodRoot As Node

    Set cProject = New clsProperties
    
    ' create root node.
    Set nodRoot = tvFavorites.Nodes.Add()
    nodRoot.Key = "Root"
    nodRoot.Text = "Favorites"
    nodRoot.Image = "Favorite"
    
    ' go read favorites text file.
    ReadFile
            
    ' expand root node.
    nodRoot.Expanded = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    
    ' go save favorites text file?
    If m_bSaveFavorites Then WriteFile
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cProject = Nothing
    
End Sub

Private Sub NewFolder_Click()

    ' create new folder node only in root node
    ' A very lazy way of doing it as it just creates a folder node
    ' called "New Folder".  If there already is a "New Folder" this
    ' must be renamed manually first.
    
    Dim nodFolder As Node
    Dim sError As String
    Const sKey As String = "New Folder"
    
    On Error GoTo ErrHandler

    ' create child folder node in root.
    Set nodFolder = tvFavorites.Nodes.Add("Root", tvwChild, sKey, sKey, "FldrClose")
    nodFolder.ExpandedImage = "FldrOpen"
        
    Exit Sub
    
ErrHandler:
    If Err.Number = 35602 Then sError = "Rename existing " & sKey
        
    MsgBox "Error:" & Err.Number & vbNewLine & Err.Description & vbNewLine & sError
    
End Sub

Private Sub OKButton_Click()

    Unload Me
    
End Sub

Private Sub tvFavorites_AfterLabelEdit(Cancel As Integer, NewString As String)

    Dim nodFolder As Node
    Dim sError As String
    
    On Error GoTo ErrHandler
    
    Set nodFolder = tvFavorites.SelectedItem
    
    ' will raise error if duplicate.
    nodFolder.Key = NewString
        
    m_bSaveFavorites = True
    
    Exit Sub
    
ErrHandler:
    Cancel = True
    If Err.Number = 35602 Then sError = "Cannot rename to " & NewString
        
    MsgBox "Error:" & Err.Number & vbNewLine & Err.Description & vbNewLine & sError
    
End Sub

Private Sub tvFavorites_BeforeLabelEdit(Cancel As Integer)
    
    ' No editing for project node.
         
    ' Cancel label edit?
    If IsProjectNode(tvFavorites.SelectedItem) Then Cancel = True
    
End Sub

Private Sub tvFavorites_DblClick()

    ' shows message box with projects properties.
    
    Dim sBuffer As String
    Dim sPath As String
    
    On Error GoTo ErrHandler
    
    sPath = tvFavorites.SelectedItem.Tag
    
    If FileExists(sPath) Then
        With cProject
            ' must initialise class with path before reading properties.
            .PathProject = sPath
            
            sBuffer = "Name: " & .Name & vbNewLine
            sBuffer = sBuffer & "Title: " & .Title & vbNewLine
            sBuffer = sBuffer & "Path: " & sPath & vbNewLine
            sBuffer = sBuffer & "Description: " & .Description & vbNewLine
            sBuffer = sBuffer & "Version: " & .Version & vbNewLine
            sBuffer = sBuffer & "Exe Type: " & .TypeExe & vbNewLine
            sBuffer = sBuffer & "Date: " & .FileDate
            
        End With
          
        MsgBox sBuffer, vbInformation, "Project Properties"
    
    End If
    
ErrHandler:
    
End Sub

Private Sub tvFavorites_DragDrop(Source As Control, x As Single, y As Single)

    ' Drop project node.
    
    On Error GoTo ErrHandler
        
    With tvFavorites
        If IsProjectNode(.DropHighlight) Then GoTo ErrHandler
        
        If m_nodX <> .DropHighlight Then
            Set m_nodX.Parent = .DropHighlight  'Set child's parent
            m_bSaveFavorites = True
                         
        End If
    End With
      
ErrHandler:
    Set tvFavorites.DropHighlight = Nothing
    Set m_nodX = Nothing
    m_bDragNode = False
    
End Sub

Private Sub tvFavorites_DragOver(Source As Control, x As Single, y As Single, State As Integer)

    If m_bDragNode Then
        ' Set DropHighlight to the mouse's coordinates.
        Set tvFavorites.DropHighlight = tvFavorites.HitTest(x, y)
        
    End If
    
End Sub

Private Sub tvFavorites_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' move a project node.  set tooltip.
    
    On Error GoTo ErrHandler
        
    Select Case Button
        Case 0
            ' reference node under cursor, not selected node.
            Set m_nodX = tvFavorites.HitTest(x, y)
              
            ' test for project node.
            If IsProjectNode(m_nodX) Then
                tvFavorites.ToolTipText = m_nodX.Tag    ' set tooltip.
                
            Else
                tvFavorites.ToolTipText = ""
                                
            End If
            
        Case vbLeftButton
            tvFavorites.ToolTipText = ""
            
            If m_bDragNode Then
                With tvFavorites
                    .DragIcon = m_nodX.CreateDragImage   ' Set the drag icon
                    .Drag vbBeginDrag ' begin drag operation.
                End With
                
            End If
            
    End Select
                            
    Exit Sub
    
ErrHandler:
    tvFavorites.Drag vbCancel
    m_bDragNode = False
    Set m_nodX = Nothing
    
End Sub

Private Sub tvFavorites_NodeClick(ByVal Node As MSComctlLib.Node)

    ' reference node.
    Set m_nodX = Node ' Set the item being dragged.
    
    ' allow node to be dragged?
    m_bDragNode = IsProjectNode(m_nodX)
    
End Sub

Private Sub tvFavorites_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    ' add a project node to highlighted node.
        
    Dim nodProject As Node
    Dim nodFolder As Node
    Dim sError As String
    Dim sKey As String
    Dim sPath As String
        
    On Error GoTo ErrHandler
        
    Set nodFolder = tvFavorites.DropHighlight
    
    ' use select node if no drophighlighted.
    If nodFolder Is Nothing Then Set nodFolder = tvFavorites.SelectedItem
               
    sPath = Data.Files(1)   'only one file is proccessed
    
    ' check the attributes of the file being dropped.
    If (GetAttr(sPath) And vbDirectory) <> vbDirectory Then
        ' check extension for correct file type.
        If GetFileExtension(sPath) = "vbp" Then
            If IsProjectNode(nodFolder) Then
                MsgBox "Cannot add a project to a project node."
                
            Else
                ' Create project node.  Raises error if duplicated anywhere in tree.
                Set nodProject = tvFavorites.Nodes.Add(nodFolder.Key, tvwChild, sPath, , "Project")
            
                ' must initialise class with path before reading properties.
                cProject.PathProject = sPath
                nodProject.Tag = sPath
                nodProject.Text = cProject.Name
                m_bSaveFavorites = True
                  
                ' make highlighted node the selected node.
                Set tvFavorites.SelectedItem = tvFavorites.DropHighlight
                
            End If
                                     
        End If
        
    End If
        
    Set tvFavorites.DropHighlight = Nothing
    
    Exit Sub
    
ErrHandler:
    Set tvFavorites.DropHighlight = Nothing
    
    If Err.Number = 35602 Then sError = "Duplicate project found."
  
    MsgBox "Error:" & Err.Number & vbNewLine & Err.Description & vbNewLine & sError
    
End Sub

Private Sub tvFavorites_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    Set tvFavorites.DropHighlight = tvFavorites.HitTest(x, y)
    
End Sub

Private Function IsProjectNode(ByRef Node As MSComctlLib.Node) As Boolean

    ' Assumes: Project nodes are the only nodes with tag value set.  If you
    ' decide to use the tag property on other nodes then this function
    ' will need revising.
    On Error GoTo ErrHandler
    
    If Len(Node.Tag) Then IsProjectNode = True
    
    Exit Function
    
ErrHandler:
   IsProjectNode = False
   
End Function

Private Sub ReadFile()

    ' read favorites file and populate tree.
    ' Does not allow duplicate project nodes.
        
    Dim nodFolder As Node
    Dim nodProject As Node
    Dim sFolder As String
    Dim sPath As String
    Dim iFile As Integer
    
    On Error Resume Next
    
    ' does favorites file exits?
    If Not FileExists(m_sPath) Then Exit Sub
    
    ' show busy pointer.
    Screen.MousePointer = vbHourglass

    iFile = FreeFile
        
    Open m_sPath For Input As #iFile    ' Open file for input.
        Do While Not EOF(iFile)         ' Loop until end of file.
            
            DoEvents
            
            Input #iFile, sFolder, sPath   ' read data into two variables.
            
            ' check existance of folder node in collection.
            For Each nodFolder In tvFavorites.Nodes
                If nodFolder.Key = sFolder Then Exit For
               
            Next
                        
            ' Create folder node?
            If nodFolder Is Nothing Then
                Set nodFolder = tvFavorites.Nodes.Add("Root", tvwChild)
                nodFolder.Key = sFolder
                nodFolder.Text = sFolder
                nodFolder.Image = "FldrClose"
                nodFolder.ExpandedImage = "FldrOpen"
               
            End If
            
            ' Create project node.  Raises error if duplicated anywhere in tree.
            Set nodProject = tvFavorites.Nodes.Add(nodFolder.Key, tvwChild, sPath, , "Project")
            
            ' Did the project node get added to the collection?
            If nodProject Is Nothing Then
                ' do nothing
                            
            Else    ' add other info
                nodProject.Tag = sPath
                
                If FileExists(sPath) Then
                    ' must initialise class with path before reading properties.
                    cProject.PathProject = sPath
                    nodProject.Text = cProject.Name
            
                Else
                    nodProject.Text = "Error:File not found"
                    
                End If
                
                Set nodProject = Nothing    ' release reference
                
            End If
                             
        Loop
        
    Close #iFile
               
    ' Return mouse pointer to normal.
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub WriteFile()

    ' Save our favorites to a file, maintain tree structure.
        
    Dim nodX As Node
    Dim sPath As String
    Dim sFolder As String
    Dim iFile As Integer
        
    On Error Resume Next
    
    ' show busy pointer.
    Screen.MousePointer = vbHourglass
    
    iFile = FreeFile
    
    Open m_sPath For Output As #iFile    ' Open file for output.
        ' check existance of folder node.
        For Each nodX In tvFavorites.Nodes
            sPath = nodX.Tag
            sFolder = nodX.Key
            
            If sFolder <> nodX.Root.Key Then sFolder = nodX.Parent.Key
                       
            If IsProjectNode(nodX) Then Write #iFile, sFolder, sPath   ' Write comma-delimited data.
        
        Next
                      
    Close #iFile    ' Close file.

    ' Return mouse pointer to normal.
    Screen.MousePointer = vbDefault
    
End Sub


