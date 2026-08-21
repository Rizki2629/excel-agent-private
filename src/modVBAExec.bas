Attribute VB_Name = "modVBAExec"
'==============================================================================
' Excel Agent Add-in - VBA Execution Engine Module
' Handles VBA code execution and module management
'==============================================================================

Option Explicit

'==============================================================================
' Constants
'==============================================================================
Private Const VBEXT_CT_StdModule As Long = 1
Private Const VBEXT_CT_ClassModule As Long = 2
Private Const VBEXT_CT_Document As Long = 100

'==============================================================================
' ExecuteResult - Structure to hold execution result
'==============================================================================
Public Type ExecuteResult
    success As Boolean
    errorMessage As String
    output As String
End Type

'==============================================================================
' ExecuteAction - Route and execute action from API response
'==============================================================================
Public Function ExecuteAction(code As String, action As String, target As String) As ExecuteResult
    On Error GoTo ErrorHandler

    Dim result As ExecuteResult

    Select Case LCase(action)
        Case "execute_vba"
            result = ExecuteVBACode(code, target)
        Case "create_module"
            Dim moduleName As String
            moduleName = GenerateModuleName()
            result = CreateNewModule(code, moduleName)
        Case "modify_code"
            ' For modify, we need to parse module name from code
            Dim targetModule As String
            targetModule = ExtractModuleName(code)
            If targetModule = "" Then targetModule = "GeneratedModule"
            result = ModifyExistingModule(code, targetModule)
        Case "run_macro"
            Dim macroName As String
            macroName = ExtractMacroName(code)
            result = RunMacro(macroName)
        Case Else
            result.success = False
            result.errorMessage = "Unknown action: " & action
    End Select

    ExecuteAction = result

    Exit Function

ErrorHandler:
    ExecuteAction.success = False
    ExecuteAction.errorMessage = "Execution Error: " & Err.Description
End Function

'==============================================================================
' ExecuteVBACode - Execute VBA code anonymously (temporary module)
'==============================================================================
Public Function ExecuteVBACode(code As String, target As String) As ExecuteResult
    On Error GoTo ErrorHandler

    Dim result As ExecuteResult
    Dim wb As Workbook
    Set wb = GetTargetWorkbook(target)

    If wb Is Nothing Then
        result.success = False
        result.errorMessage = "Target workbook not found"
        ExecuteVBACode = result
        Exit Function
    End If

    ' Create temporary module
    Dim vbComp As Object
    Set vbComp = wb.VBProject.VBComponents.Add(VBEXT_CT_StdModule)

    ' Add code to module
    vbComp.CodeModule.AddFromString code

    ' Get the first procedure name
    Dim procName As String
    procName = GetFirstProcedureName(vbComp)

    If procName = "" Then
        result.success = False
        result.errorMessage = "No procedure found in code"
        GoTo Cleanup
    End If

    ' Execute the procedure
    Application.Run procName

    result.success = True
    result.output = "Executed: " & procName

Cleanup:
    ' Remove temporary module
    On Error Resume Next
    wb.VBProject.VBComponents.Remove vbComp
    On Error GoTo ErrorHandler

    ExecuteVBACode = result

    Exit Function

ErrorHandler:
    result.success = False
    result.errorMessage = "Error executing VBA: " & Err.Description
    ExecuteVBACode = result
End Function

'==============================================================================
' CreateNewModule - Create a new named VBA module
'==============================================================================
Public Function CreateNewModule(code As String, moduleName As String) As ExecuteResult
    On Error GoTo ErrorHandler

    Dim result As ExecuteResult
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        result.success = False
        result.errorMessage = "No active workbook"
        CreateNewModule = result
        Exit Function
    End If

    ' Remove existing module with same name
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(moduleName)
    On Error GoTo ErrorHandler

    ' Create new module
    Dim vbComp As Object
    Set vbComp = wb.VBProject.VBComponents.Add(VBEXT_CT_StdModule)
    vbComp.Name = moduleName

    ' Add code
    vbComp.CodeModule.AddFromString code

    result.success = True
    result.output = "Created module: " & moduleName

    CreateNewModule = result

    Exit Function

ErrorHandler:
    result.success = False
    result.errorMessage = "Error creating module: " & Err.Description
    CreateNewModule = result
End Function

'==============================================================================
' ModifyExistingModule - Modify an existing VBA module
'==============================================================================
Public Function ModifyExistingModule(code As String, moduleName As String) As ExecuteResult
    On Error GoTo ErrorHandler

    Dim result As ExecuteResult
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        result.success = False
        result.errorMessage = "No active workbook"
        ModifyExistingModule = result
        Exit Function
    End If

    ' Check if module exists
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = wb.VBProject.VBComponents(moduleName)
    On Error GoTo ErrorHandler

    If vbComp Is Nothing Then
        ' Module doesn't exist, create it
        result = CreateNewModule(code, moduleName)
        ModifyExistingModule = result
        Exit Function
    End If

    ' Clear existing code
    If vbComp.CodeModule.CountOfLines > 0 Then
        vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
    End If

    ' Add new code
    vbComp.CodeModule.AddFromString code

    result.success = True
    result.output = "Modified module: " & moduleName

    ModifyExistingModule = result

    Exit Function

ErrorHandler:
    result.success = False
    result.errorMessage = "Error modifying module: " & Err.Description
    ModifyExistingModule = result
End Function

'==============================================================================
' RunMacro - Run an existing macro by name
'==============================================================================
Public Function RunMacro(macroName As String) As ExecuteResult
    On Error GoTo ErrorHandler

    Dim result As ExecuteResult

    Application.Run macroName

    result.success = True
    result.output = "Executed macro: " & macroName

    RunMacro = result

    Exit Function

ErrorHandler:
    result.success = False
    result.errorMessage = "Error running macro: " & Err.Description
    RunMacro = result
End Function

'==============================================================================
' GetTargetWorkbook - Get workbook by target specification
'==============================================================================
Private Function GetTargetWorkbook(target As String) As Workbook
    Select Case LCase(target)
        Case "active_workbook", "active", ""
            Set GetTargetWorkbook = ActiveWorkbook
        Case Else
            ' Try to find workbook by name
            On Error Resume Next
            Set GetTargetWorkbook = Workbooks(target)
            On Error GoTo 0
    End Select
End Function

'==============================================================================
' GetFirstProcedureName - Get name of first procedure in module
'==============================================================================
Private Function GetFirstProcedureName(vbComp As Object) As String
    Dim i As Long
    Dim procName As String
    Dim lineText As String

    For i = 1 To vbComp.CodeModule.CountOfLines
        lineText = Trim(vbComp.CodeModule.Lines(i, 1))

        ' Look for Sub or Function declaration
        If Left(LCase(lineText), 4) = "sub " Or Left(LCase(lineText), 9) = "function " Then
            ' Extract procedure name
            If Left(LCase(lineText), 4) = "sub " Then
                procName = Mid(lineText, 5)
            Else
                procName = Mid(lineText, 10)
            End If

            ' Remove parentheses and parameters
            Dim parenPos As Long
            parenPos = InStr(procName, "(")
            If parenPos > 0 Then
                procName = Left(procName, parenPos - 1)
            End If

            ' Remove attribute line if present
            If Left(LCase(Trim(procName)), 10) = "attribute " Then
                ' Skip attribute lines
            Else
                GetFirstProcedureName = Trim(procName)
                Exit Function
            End If
        End If
    Next i

    GetFirstProcedureName = ""
End Function

'==============================================================================
' GenerateModuleName - Generate a unique module name
'==============================================================================
Private Function GenerateModuleName() As String
    Dim timestamp As String
    timestamp = Format(Now, "yyyymmddhhmmss")

    GenerateModuleName = "AgentModule_" & timestamp
End Function

'==============================================================================
' ExtractModuleName - Extract module name from code comment
'==============================================================================
Private Function ExtractModuleName(code As String) As String
    ' Look for pattern: ' Module: ModuleName
    Dim lines() As String
    lines = Split(code, vbCrLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim(lines(i))

        If Left(LCase(lineText), 9) = "' module:" Then
            ExtractModuleName = Trim(Mid(lineText, 10))
            Exit Function
        End If
    Next i

    ExtractModuleName = ""
End Function

'==============================================================================
' ExtractMacroName - Extract macro name from code
'==============================================================================
Private Function ExtractMacroName(code As String) As String
    ' Look for pattern: Sub MacroName() or Call MacroName
    Dim lines() As String
    lines = Split(code, vbCrLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim(lines(i))

        ' Check for Sub declaration
        If Left(LCase(lineText), 4) = "sub " Then
            Dim procName As String
            procName = Mid(lineText, 5)

            Dim parenPos As Long
            parenPos = InStr(procName, "(")
            If parenPos > 0 Then
                procName = Left(procName, parenPos - 1)
            End If

            ExtractMacroName = Trim(procName)
            Exit Function
        End If

        ' Check for Call statement
        If Left(LCase(lineText), 5) = "call " Then
            ExtractMacroName = Trim(Mid(lineText, 6))
            Exit Function
        End If
    Next i

    ExtractMacroName = ""
End Function

'==============================================================================
' ListModules - List all VBA modules in active workbook
'==============================================================================
Public Function ListModules() As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        ListModules = "No active workbook"
        Exit Function
    End If

    Dim result As String
    Dim vbComp As Object

    For Each vbComp In wb.VBProject.VBComponents
        result = result & vbComp.Name & " (" & GetComponentType(vbComp.Type) & ")" & vbCrLf
    Next vbComp

    If result = "" Then
        result = "No modules found"
    End If

    ListModules = result
End Function

'==============================================================================
' GetComponentType - Get component type name
'==============================================================================
Private Function GetComponentType(compType As Long) As String
    Select Case compType
        Case VBEXT_CT_StdModule
            GetComponentType = "Module"
        Case VBEXT_CT_ClassModule
            GetComponentType = "Class"
        Case VBEXT_CT_Document
            GetComponentType = "Document"
        Case Else
            GetComponentType = "Unknown"
    End Select
End Function

'==============================================================================
' DeleteModule - Delete a VBA module by name
'==============================================================================
Public Function DeleteModule(moduleName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        DeleteModule = False
        Exit Function
    End If

    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(moduleName)

    DeleteModule = True

    Exit Function

ErrorHandler:
    DeleteModule = False
End Function
