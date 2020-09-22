VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ProjectProps
    ppName As String
    ppValue As String
    
End Type

Private PropList() As ProjectProps

'local vars
Private m_sDescription As String
Private m_sName As String
Private m_sPathProject As String
Private m_sTitle As String
Private m_sTypeExe As String
Private m_sVersion As String
Private m_sDate As Date

Public Property Get Description() As String

    Description = m_sDescription
    
End Property

Public Property Get FileDate() As Date

    FileDate = m_sDate
    
End Property

Public Property Get Name() As String

    ' project name
    Name = m_sName
    
End Property

Public Property Get Title() As String

    Title = m_sTitle
    
End Property

Public Property Get TypeExe() As String

    TypeExe = m_sTypeExe
    
End Property

Public Property Get Version() As String
    
    ' project version   format: major, minor, revision
    Version = m_sVersion
    
End Property

Public Property Let PathProject(ByVal sPath As String)

    ' no attempt is made to check if the project file exists, this is
    ' done by calling procedure.  Although class will not produce
    ' fatal error if project file does not exists.
    
    On Error GoTo ErrHandler
    
    If sPath <> m_sPathProject Then
        m_sPathProject = sPath
        
        ' must call procedure that populates udt from the project file first.
        ReadProject m_sPathProject
        
        ' initialise class properties.
        InitProps
        
    End If
    
ErrHandler:

End Property

Public Property Get PathProject() As String
    
    ' retuns absolute path, this read property is not used.
    PathProject = m_sPathProject
        
End Property

Private Sub InitProps()

    ' With the aid of a couple of functions assign values to classes read only properties
    ' GetValue function ensures all properties are assigned or cleared!
    
    Dim sTemp As String
                
    On Error Resume Next
    
    m_sDate = 0
    
    ' project name
    sTemp = GetValue("name")    ' getvalue expects lower case.
    m_sName = TrimQuotes(sTemp)
    
    ' title
    sTemp = GetValue("title")
    m_sTitle = TrimQuotes(sTemp)
    
    ' version
    m_sVersion = GetValue("majorver") & "," & GetValue("minorver") & "," & GetValue("revisionver")
           
    ' description
    sTemp = GetValue("description")
    m_sDescription = TrimQuotes(sTemp)
       
    ' type
    sTemp = GetValue("type")
    m_sTypeExe = TrimQuotes(sTemp)
     
    ' date
    m_sDate = FileDateTime(m_sPathProject)  ' will raise error here if no file
  
    ' release some memory
    Erase PropList
    
End Sub

Private Function GetValue(PropName As String) As String

    ' returns the value for a given name from the udt array.
    ' Assumes lowercase input parameter
       
    Dim i As Integer
    Dim sBuffer As String
    
    On Error GoTo ErrHandler
    
    For i = 0 To UBound(PropList)
        With PropList(i)
            If .ppName = PropName Then
                sBuffer = .ppValue
                Exit For
                
            End If
        End With
        
    Next
    
    GetValue = sBuffer
    
    Exit Function
    
ErrHandler:
    GetValue = ""
    
End Function

Private Sub ReadProject(PathToProject As String)

    'read project file, initialise udt
    
    Dim iPos As Integer
    Dim iIndex As Integer
    Dim nFile As Integer
    Dim sTextLine As String
    Dim sLH As String
    Dim sRH As String
           
    On Error GoTo ErrHandler
        
    nFile = FreeFile
    
    Open PathToProject For Input As #nFile   ' Open file.
        Do While Not EOF(nFile)   ' Loop until end of file.
            Line Input #nFile, sTextLine   ' Read line into variable.
            
            BreakString sTextLine, sLH, sRH, &H3D   ' break char "="

            If Len(sRH) > 0 Then
                ReDim Preserve PropList(iIndex)
                
                With PropList(iIndex)
                    .ppName = LCase$(sLH)
                    .ppValue = sRH
                    
                End With
                
                iIndex = iIndex + 1 ' increment
                
            End If
        
        Loop
    Close #nFile   ' Close file.

    Exit Sub
    
ErrHandler:
    Close ' close any files left open.
    Debug.Print "Error:" & Err.Number & vbNewLine & Err.Description
    
End Sub
