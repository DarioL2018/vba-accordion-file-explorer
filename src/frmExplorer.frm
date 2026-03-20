VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExplorer 
   Caption         =   "Template Explorer"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "frmExplorer.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- Windows API for INI Reading ---
#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If

Dim RootPath As String
Dim ExpandedFolders As Object ' Dictionary to remember which folders are open

' --- Form Initialization ---
Private Sub UserForm_Initialize()
    Set ExpandedFolders = CreateObject("Scripting.Dictionary")
    ExpandedFolders.CompareMode = 1 ' TextCompare (case-insensitive)
    
    Dim IniPath As String
    IniPath = ThisDocument.Path & "\config.ini"
    
    ' 1. Try to read from INI first (Secret backdoor for future changes)
    RootPath = ReadConfig("Settings", "RootFolder", IniPath)
    
    ' 2. If no INI exists or it's empty, use the client's hardcoded default path
    If RootPath = "" Then
        RootPath = "C:\Users\mborrelli\Documents\Neo Iustec\Vorlagen"
    End If
    
    ' 3. Safety fallback: if the hardcoded path doesn't exist on the current PC (e.g., your testing PC)
    If Dir(RootPath, vbDirectory) = "" Then
        RootPath = ThisDocument.Path
    End If
    
    txtPath.Text = RootPath
    RefreshTree
End Sub

Private Function ReadConfig(Section As String, Key As String, FilePath As String) As String
    Dim Result As String * 255
    Dim Length As Long
    Length = GetPrivateProfileString(Section, Key, "", Result, 255, FilePath)
    If Length > 0 Then ReadConfig = Left(Result, Length) Else ReadConfig = ""
End Function

' --- Core Tree Logic ---
Private Sub RefreshTree()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    lstFiles.Clear
    
    If FSO.FolderExists(RootPath) Then
        LoadNode RootPath, 0, FSO
    End If
    
    Set FSO = Nothing
End Sub

' Recursive function to build the tree (With Permission Protection)
Private Sub LoadNode(ByVal Path As String, ByVal IndentLevel As Integer, FSO As Object)
    Dim Folder As Object, SubFolder As Object, File As Object
    
    ' 0. Attempt to access the folder. If Windows blocks access (permissions), exit silently.
    On Error Resume Next
    Set Folder = FSO.GetFolder(Path)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim IndentSpace As String
    IndentSpace = String(IndentLevel * 4, " ") ' 4 spaces per depth level
    
    ' 1. Load Subfolders
    On Error Resume Next
    For Each SubFolder In Folder.SubFolders
        If Err.Number = 0 Then ' Only process if Windows grants permission
            Dim IsExpanded As Boolean
            IsExpanded = ExpandedFolders.Exists(SubFolder.Path)
            
            ' Add the folder to the list simulating the TreeView
            If IsExpanded Then
                lstFiles.AddItem IndentSpace & "v   " & SubFolder.Name
            Else
                lstFiles.AddItem IndentSpace & ">   " & SubFolder.Name
            End If
            
            ' Store data in hidden columns
            lstFiles.List(lstFiles.ListCount - 1, 1) = SubFolder.Path
            lstFiles.List(lstFiles.ListCount - 1, 2) = "FOLDER"
            
            ' Recursion
            If IsExpanded Then
                LoadNode SubFolder.Path, IndentLevel + 1, FSO
            End If
        End If
        Err.Clear ' Clear error in case the next folder is blocked
    Next SubFolder
    On Error GoTo 0
    
    ' 2. Load Word Files
    On Error Resume Next
    For Each File In Folder.Files
        If Err.Number = 0 Then
            ' Filter for Word extensions and ignore temp files (~$)
            If (InStr(1, File.Name, ".doc") > 0 Or InStr(1, File.Name, ".dot") > 0) And Left(File.Name, 2) <> "~$" Then
                lstFiles.AddItem IndentSpace & "       " & File.Name
                lstFiles.List(lstFiles.ListCount - 1, 1) = File.Path
                lstFiles.List(lstFiles.ListCount - 1, 2) = "FILE"
            End If
        End If
        Err.Clear
    Next File
    On Error GoTo 0
End Sub

' --- User Interactions ---

' Single Click: Expand / Collapse folders
Private Sub lstFiles_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Or lstFiles.ListIndex = -1 Then Exit Sub ' Left click only
    
    Dim ItemPath As String
    Dim ItemType As String
    
    ItemPath = lstFiles.List(lstFiles.ListIndex, 1)
    ItemType = lstFiles.List(lstFiles.ListIndex, 2)
    
    If ItemType = "FOLDER" Then
        ' Toggle state
        If ExpandedFolders.Exists(ItemPath) Then
            ExpandedFolders.Remove ItemPath
        Else
            ExpandedFolders.Add ItemPath, True
        End If
        RefreshTree ' Redraw the tree
    End If
End Sub

' Double Click: Open files
Private Sub lstFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstFiles.ListIndex = -1 Then Exit Sub
    
    Dim ItemPath As String
    Dim ItemType As String
    
    ItemPath = lstFiles.List(lstFiles.ListIndex, 1)
    ItemType = lstFiles.List(lstFiles.ListIndex, 2)
    
    If ItemType = "FILE" Then
        OpenSelectedFile ItemPath
    End If
End Sub

' Open Button
Private Sub btnOpen_Click()
    If lstFiles.ListIndex = -1 Then Exit Sub
    
    If lstFiles.List(lstFiles.ListIndex, 2) = "FILE" Then
        OpenSelectedFile lstFiles.List(lstFiles.ListIndex, 1)
    End If
End Sub

' Cancel Button
Private Sub btnCancel_Click()
    Unload Me
End Sub

' Helper function to open files
Private Sub OpenSelectedFile(ByVal FullPath As String)
    On Error GoTo ErrorHandler
    
    If InStr(1, FullPath, ".dot") > 0 Then
        Documents.Add Template:=FullPath ' Create new document from template
    Else
        Documents.Open FileName:=FullPath ' Open normal document
    End If
    
    'Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Error opening file: " & Err.Description, vbExclamation, "System Error"
End Sub

