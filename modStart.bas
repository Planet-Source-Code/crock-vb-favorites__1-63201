Attribute VB_Name = "modStart"
Option Explicit

Private Enum ShowWindowType
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MINIMIZED = 2
    SW_MAXIMIZED = 3
End Enum

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As ShowWindowType) As Long

Public fMainForm As frmMain

Sub Main()

    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Show
    
End Sub

Public Sub RunAssociated(sFileName As String, Optional sParams As String = "", Optional sDefaultDir As String = "")
  
    Dim sBuffer As String
    Dim lrc As Long
  
    On Error Resume Next
      
    'Parameters passed to ShellExecute
    'hWnd - active form
    'lpOperation - "Open" or "Print" (vbNullString defaults to "Open"
    'lpFile - Program name or name of a for to print or open using the associated program
    'lpParameters - Command line if lpFile is a program to run
    'lpDirectory - Default directory to use
    'nShowCmd - Constant specifying how to show the launched program (maximized, minimized, normal)
    
    lrc = ShellExecute(0&, "Open", sFileName, sParams, sDefaultDir, SW_NORMAL)
  
    'If the function fails, the return value is an error value that is <= to 32.
    If lrc <= 32 Then
        sBuffer = "ShellExecute Error:" & lrc & vbNewLine
        sBuffer = sBuffer & ShellError(CInt(lrc)) & vbNewLine & sFileName
        MsgBox sBuffer, vbCritical
        
    End If
 
End Sub

Private Function ShellError(ErrNumber As Integer) As String

    Dim sError As String
    
    Select Case ErrNumber
        Case 0
            sError = "The operating system is out of memory or resources."
        Case 1
            sError = "The specified file was not found."
        Case 2
            sError = "The specified path was not found."
        Case 3
            sError = "The .EXE file is invalid (non-Win32 .EXE or error in .EXE image)."
        Case 4
            sError = "The operating system denied access to the specified file."
        Case 5
            sError = "The filename association is incomplete or invalid."
        Case 6
            sError = "The DDE transaction could not be completed because other DDE transactions were being processed."
        Case 7
            sError = "The DDE transaction failed."
        Case 8
            sError = "The DDE transaction could not be completed because the request timed out."
        Case 9
            sError = "The specified dynamic-link library was not found."
        Case 10
            sError = "The specified file was not found."
        Case 11
            sError = "There is no application associated with the given filename extension."
        Case 12
            sError = "There was not enough memory to complete the operation."
        Case 13
            sError = "The specified path was not found."
        Case 14
            sError = "A sharing violation occurred."
        Case Else
            sError = "Unknown Error!"
    End Select

    ShellError = sError
    
End Function

Public Function FileExists(ByVal sPath) As Boolean

    On Error GoTo ErrHandler
    
    If FileLen(sPath) >= 0 Then FileExists = True
     
    Exit Function
       
ErrHandler:
    FileExists = False
    
End Function

Public Function TrimQuotes(ByVal sInput As String) As String

    'Trim Quotes
    Dim ilength As Integer
    
    On Error GoTo ErrHandler
    
    ilength = Len(sInput)
    
    If Right$(sInput, 1) = Chr$(34) Then ilength = ilength - 1
    
    If Left$(sInput, 1) = Chr$(34) Then
        ilength = ilength - 1
        TrimQuotes = Mid$(sInput, 2, ilength)
        
    Else
        TrimQuotes = Mid$(sInput, 1, ilength)
        
    End If
       
    Exit Function
       
ErrHandler:
    TrimQuotes = ""
    
End Function

Public Sub BreakString(ByVal StringToBreak As String, ByRef StringA As String, ByRef StringB As String, ByVal BreakChar As Byte)

    ' Break the string at the first occurance of the break char
    ' be aware that the stringa and stringb are passed by reference.
    
    Dim iPos As String
    
    On Error Resume Next
    
    StringA = ""
    StringB = ""
    
    iPos = InStr(1, StringToBreak, Chr$(BreakChar), vbBinaryCompare)
    
    If iPos Then
        StringA = Left$(StringToBreak, iPos - 1)
        StringA = Trim$(StringA)
        StringB = Trim$(Mid$(StringToBreak, iPos + 1))
        
    End If
    
End Sub

Public Function GetFileExtension(ByVal sPath As String) As String

    ' Return lowercase file extension.
    
    Dim i As Integer
    Dim sBuffer As String
    
    ' Get extension.
    For i = Len(sPath) To 1 Step -1
        If Mid$(sPath, i, 1) = "." Then
            sBuffer = Mid$(sPath, i + 1)
            Exit For
            
        End If
        
    Next
    
    GetFileExtension = LCase$(sBuffer)
    
End Function
