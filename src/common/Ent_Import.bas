Attribute VB_Name = "Ent_Import"
Option Explicit

' ============================================
' Module   : Ent_Import
' Layer    : Entry
' Purpose  : Import VBA modules from files
'            - Bulk import bas/cls/frm files
'            - Overwrite existing modules with same name
' Version  : 1.0.1
' Created  : 2026-01-30
' Updated  : 2026-01-30 - Use late binding (no VBIDE reference required)
' ============================================

' Component Type constants (for late binding)
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Private m_ImportCount As Long
Private m_SkipCount As Long
Private m_ErrorCount As Long

' ============================================
' Import_AllModulesEx
' Import all modules from separate folders for each type
' ============================================
Public Sub Import_AllModulesEx( _
    Optional ByVal basDir As String = "", _
    Optional ByVal clsDir As String = "", _
    Optional ByVal frmDir As String = "", _
    Optional ByVal overwriteExisting As Boolean = True)

    Dim vbProj As Object  ' VBIDE.VBProject (late binding)
    Set vbProj = Application.VBE.ActiveVBProject

    Dim defaultBase As String
    defaultBase = ThisWorkbook.Path & "\src"

    ' Set defaults
    If basDir = "" Then basDir = defaultBase & "\bas"
    If clsDir = "" Then clsDir = defaultBase & "\cls"
    If frmDir = "" Then frmDir = defaultBase & "\frm"

    ' Reset counters
    m_ImportCount = 0
    m_SkipCount = 0
    m_ErrorCount = 0

    ' Import from each folder
    ImportFromFolder vbProj, basDir, "*.bas", overwriteExisting
    ImportFromFolder vbProj, clsDir, "*.cls", overwriteExisting
    ImportFromFolder vbProj, frmDir, "*.frm", overwriteExisting

    MsgBox "Import completed:" & vbCrLf & _
           "Imported: " & m_ImportCount & vbCrLf & _
           "Skipped: " & m_SkipCount & vbCrLf & _
           "Errors: " & m_ErrorCount, vbInformation, "Import Result"
End Sub

' ============================================
' Import_AllModules
' Import all modules from a base folder with subfolders
' ============================================
Public Sub Import_AllModules( _
    Optional ByVal baseDir As String = "", _
    Optional ByVal overwriteExisting As Boolean = True)

    If baseDir = "" Then
        baseDir = ThisWorkbook.Path & "\src"
    End If

    Import_AllModulesEx _
        basDir:=baseDir & "\bas", _
        clsDir:=baseDir & "\cls", _
        frmDir:=baseDir & "\frm", _
        overwriteExisting:=overwriteExisting
End Sub

' ============================================
' Import_SingleModule
' Import a single module file
' ============================================
Public Function Import_SingleModule( _
    ByVal filePath As String, _
    Optional ByVal overwriteExisting As Boolean = True) As Boolean

    On Error GoTo EH

    Dim vbProj As Object  ' VBIDE.VBProject (late binding)
    Set vbProj = Application.VBE.ActiveVBProject

    Import_SingleModule = ImportModuleFile(vbProj, filePath, overwriteExisting)
    Exit Function

EH:
    Debug.Print "Import_SingleModule error: " & filePath & " - " & Err.Description
    Import_SingleModule = False
End Function

' ============================================
' Import_FromFileDialog
' Open file dialog to select and import modules
' ============================================
Public Sub Import_FromFileDialog()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select VBA Modules to Import"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "VBA Modules", "*.bas;*.cls;*.frm"
        .Filters.Add "Standard Modules", "*.bas"
        .Filters.Add "Class Modules", "*.cls"
        .Filters.Add "UserForms", "*.frm"
        .Filters.Add "All Files", "*.*"

        If .Show = -1 Then
            Dim vbProj As Object  ' VBIDE.VBProject (late binding)
            Set vbProj = Application.VBE.ActiveVBProject

            m_ImportCount = 0
            m_SkipCount = 0
            m_ErrorCount = 0

            Dim i As Long
            For i = 1 To .SelectedItems.Count
                ImportModuleFile vbProj, .SelectedItems(i), True
            Next i

            MsgBox "Import completed:" & vbCrLf & _
                   "Imported: " & m_ImportCount & vbCrLf & _
                   "Skipped: " & m_SkipCount & vbCrLf & _
                   "Errors: " & m_ErrorCount, vbInformation, "Import Result"
        End If
    End With
End Sub

' ============================================
' Private Helper Methods
' ============================================

Private Sub ImportFromFolder( _
    ByVal vbProj As Object, _
    ByVal folderPath As String, _
    ByVal pattern As String, _
    ByVal overwriteExisting As Boolean)

    ' Check if folder exists
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        Debug.Print "Folder not found: " & folderPath
        Exit Sub
    End If

    ' Find files matching pattern
    Dim fileName As String
    fileName = Dir(folderPath & "\" & pattern)

    Do While Len(fileName) > 0
        Dim filePath As String
        filePath = folderPath & "\" & fileName

        ImportModuleFile vbProj, filePath, overwriteExisting

        fileName = Dir()
    Loop
End Sub

Private Function ImportModuleFile( _
    ByVal vbProj As Object, _
    ByVal filePath As String, _
    ByVal overwriteExisting As Boolean) As Boolean

    On Error GoTo EH

    ' Get module name from file
    Dim moduleName As String
    moduleName = GetModuleNameFromFile(filePath)

    If Len(moduleName) = 0 Then
        Debug.Print "Could not determine module name: " & filePath
        m_ErrorCount = m_ErrorCount + 1
        ImportModuleFile = False
        Exit Function
    End If

    ' Check if module already exists
    Dim existingComp As Object  ' VBIDE.VBComponent (late binding)
    Set existingComp = FindComponent(vbProj, moduleName)

    If Not existingComp Is Nothing Then
        If overwriteExisting Then
            ' Remove existing module
            If CanRemoveComponent(existingComp) Then
                vbProj.VBComponents.Remove existingComp
                Debug.Print "Removed existing: " & moduleName
            Else
                Debug.Print "Cannot remove (Document module): " & moduleName
                m_SkipCount = m_SkipCount + 1
                ImportModuleFile = False
                Exit Function
            End If
        Else
            Debug.Print "Skipped (already exists): " & moduleName
            m_SkipCount = m_SkipCount + 1
            ImportModuleFile = False
            Exit Function
        End If
    End If

    ' Import the module
    Dim imported As Object  ' VBIDE.VBComponent
    Set imported = vbProj.VBComponents.Import(filePath)

    ' Rename if the imported name doesn't match the expected name
    If imported.Name <> moduleName Then
        imported.Name = moduleName
        Debug.Print "Imported & renamed: " & moduleName
    Else
        Debug.Print "Imported: " & moduleName
    End If

    m_ImportCount = m_ImportCount + 1
    ImportModuleFile = True
    Exit Function

EH:
    Debug.Print "Import error: " & filePath & " - " & Err.Description
    m_ErrorCount = m_ErrorCount + 1
    ImportModuleFile = False
End Function

Private Function GetModuleNameFromFile(ByVal filePath As String) As String
    ' Derive module name from filename (e.g. "Utl_File.bas" -> "Utl_File")
    ' This is reliable regardless of whether Attribute VB_Name exists in the file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetModuleNameFromFile = fso.GetBaseName(filePath)
End Function

Private Function FindComponent( _
    ByVal vbProj As Object, _
    ByVal componentName As String) As Object  ' Returns VBIDE.VBComponent or Nothing

    On Error Resume Next
    Set FindComponent = vbProj.VBComponents(componentName)
    On Error GoTo 0
End Function

Private Function CanRemoveComponent(ByVal comp As Object) As Boolean
    ' Document modules (ThisWorkbook, Sheet1, etc.) cannot be removed
    CanRemoveComponent = (comp.Type <> vbext_ct_Document)
End Function

' ============================================
' Test/Demo Procedures
' ============================================

' Test: Import from default locations
Private Sub RunImport_AllModules()
    Import_AllModules "C:\Dev\UniversalModelForVBA\Github\src", True
End Sub

' Test: Import from separate folders
Private Sub RunImport_AllModulesEx()
    Import_AllModulesEx _
        basDir:="C:\Dev\UniversalModelForVBA\Github\src\bas", _
        clsDir:="C:\Dev\UniversalModelForVBA\Github\src\cls", _
        frmDir:="C:\Dev\UniversalModelForVBA\Github\src\frm", _
        overwriteExisting:=True
End Sub

' Test: Import single file
Private Sub RunImport_SingleModule()
    Dim result As Boolean
    result = Import_SingleModule("C:\Dev\UniversalModelForVBA\Github\src\bas\Utl_Normalize.bas", True)
    MsgBox "Import result: " & result
End Sub

' Test: Import via file dialog
Private Sub RunImport_FromFileDialog()
    Import_FromFileDialog
End Sub
