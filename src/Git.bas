Attribute VB_Name = "Git"
''
' This module is created to easily export and import VBA components
' into and from `./src` directory and run basic git commands.
'
' This is needed as git can't read Excel files directly, but can read the
' source files.
'
' The export should be ran either on every save or on certain events.
'
' @author Robert Todar <robert@roberttodar.com>
' @ref Microsoft Visual Basic for Application Extensibility 5.3
''
Option Explicit

' Directory where all source code and git is stored. `./src`
Private Property Get sourceDirectory() As String
    sourceDirectory = ThisWorkbook.Path & "\src\"
End Property

' This Projects VB Components.
' @NOTE: Should this be a single project, or should I use this
'        for any project/workbook? For now will leave as the
'        current
Private Property Get components() As VBComponents
    Set components = ThisWorkbook.VBProject.VBComponents
End Property

' Helper function to run scripts from the root directory.
Private Function bash(script As String, Optional keepCommandWindowOpen As Boolean = False) As Double
    ' cmd.exe Opens the command prompt.
    ' /S      Modifies the treatment of string after /C or /K (see below)
    ' /C      Carries out the command specified by string and then terminates
    ' /K      Carries out the command specified by string but remains
    ' cd      Change directory to the root directory.
    bash = Shell("cmd.exe /S /" & IIf(keepCommandWindowOpen, "K", "C") & " cd " & ThisWorkbook.Path & " && " & script)
End Function

' Get the file name for a VBComponent. That is the component name and the proper extension.
Public Function useFileName(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            useFileName = component.Name & ".cls"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            useFileName = component.Name & ".bas"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            useFileName = component.Name & ".frm"
            
        Case vbext_ComponentType.vbext_ct_Document
            useFileName = component.Name & ".cls"
            
        Case Else
            ' @TODO: Need to think of possible throwing an error?
            ' Is it possible to get something else?? I don't think so
            ' Will need to double check this.
            Debug.Print "Unknown componant"
    End Select
End Function

' Check to see if component exits in this current Project
Private Function componentExists(ByVal fileName As String, Optional ByRef outComponent As VBComponent) As Boolean
    ' Loop each component and do something with the
    Dim index As Long
    For index = 1 To components.Count
        Dim component As VBComponent
        Set component = components(index)
        
        If useFileName(component) = fileName Then
            componentExists = True
            Set outComponent = component
            Exit Function
        End If
    Next index
End Function

' Export all modules in this current workbook into a src dir
Public Sub exportComponents()
    ' Make sure the source directory exists.
    Dir sourceDirectory, vbNormal
    
    ' Loop each component and do something with the
    Dim index As Long
    For index = 1 To components.Count
        Dim component As VBComponent
        Set component = components(index)
        
        component.Export sourceDirectory & useFileName(component)
    Next index
End Sub

' Import source code from the source Directory.
' Danger: This will cause files to overwrite that already exists.
' Daner: This will also remove files not found in the source component.
Private Sub dangerouslyImportComponents()
    Dim fso As New Scripting.FileSystemObject
    Dim fileName As String
    fileName = fso.GetFileName(FILEPATH)
    
    Dim component As VBComponent
    If componentExists(fileName, component) Then
        ' Need to change the name as it exists in memory still
        ' even after removing. It will no longer exist once code
        ' is finished executing.
        component.Name = component.Name & "_WillBeRemovedAfterExecution_"
        
        ' This removes components, but doesn't from memory until
        ' after all code execution has completed. See note above.
        components.Remove component
    End If
    
    components.Import FILEPATH
End Sub

' Commit files to local Git.
Private Sub commitToLocal()
    bash "git init && git add . && git commit -m ""This is only a test"""
    'Shell "cmd.exe /S /K cd " & ThisWorkbook.Path & "\src\ && git init && git add . && git commit -m ""This is only a test"""
End Sub

Private Function gitVersionNumber()
    Debug.Print bash("git --version", True)
End Function

Public Sub testUpload()
    ' Get the components collection
    Dim components As VBComponents
    Set components = ThisWorkbook.VBProject.VBComponents
    
    Dim fso As New Scripting.FileSystemObject
    Dim fileName As String
    fileName = fso.GetFileName(FILEPATH)
    
    Dim component As VBComponent
    If componentExists(fileName, component) Then
        component.Name = component.Name & "REMOVEME"
        components.Remove component
    End If
    
    components.Import FILEPATH
End Sub

' Converts the component enum to a string.
Private Function useComponentTypeName(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            useComponentTypeName = "Class Module"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            useComponentTypeName = "Module"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            useComponentTypeName = "Form"
            
        Case vbext_ComponentType.vbext_ct_Document
            useComponentTypeName = "Document"
            
        Case Else
            Debug.Print "Unknown"
    End Select
End Function

' Prints out details about a specific VBComponent
Private Sub printComponentDetails(ByRef component As VBComponent)
    Debug.Print component.Name, useComponentTypeName(component), useFileName(component)
End Sub

' Prints out details about all VBComponents in the current project
Private Sub printCurrentProjectComponentsDetails()
    Dim index As Long
    For index = 1 To components.Count
        Dim component As VBComponent
        Set component = components(index)
        
        printComponentDetails component
    Next index
End Sub

' Prints out details about all VBComponents in the current project
Private Sub printDiffFromSourceFolder()
    Dim index As Long
    For index = 1 To components.Count
        Dim component As VBComponent
        Set component = components(index)
        
        Debug.Print useFileName(component)
    Next index
End Sub
