Attribute VB_Name = "modModuleUpdating"
'
' !!!! work-in-progress !!!!
' created by Chris Staines
'
' interact with modules within the vba ide
' general idea is to update the modules from a remote source so that workbooks don't have to be redeployed for code updates
'
' required references:
'   Microsoft Visual Basic for Applications Extensibility 5.3 (vbe6ext.olb)
'
' notes:
'
'   to remove a module:
'       vbproject.vbcomponents.remove <module>
'
'   to import a module:
'       vbproject.vbcomponents.import <filename>
'
'   to delete all lines in a module:
'       vbcomponent.codemodule.deletelines 1, vbcomponent.codemodule.countoflines
'
'   to import lines from a file into a module:
'       vbcomponent.codemodule.addfromfile <filename>
'
'   to count lines:
'       module_acccuratecountoflines(application.vbe.vbprojects(1).vbcomponents("nameofmodule"))
'

Function Module_Overwrite(strModule As String, strFile As String) As Boolean
' clear existing code in a module and import code from a file in its place
'
' example:  Module_Overwrite("modServiceBench_v5", "\\msfs13.lowes.com\DATA1\HOME\cstaines\code\modServiceBench_v5.bas")

Dim varIsModule
' variable for results of ismodule

varIsModule = IsModule(strModule$)
' attempt to find module by name

If varIsModule(0) = True Then
' if the module was found, ...

    If Dir(strFile$) = Right(strFile$, Len(strFile$) - InStrRev(strFile$, "\")) Then
    ' if the given file exists, ...
        
        Dim vbcModule As VBComponent
        ' variable for module returned
        
        Set vbcModule = varIsModule(2)
        ' set module as module found by name
    
        vbcModule.CodeModule.DeleteLines 1, vbcModule.CodeModule.CountOfLines
        ' clear all existing code of the module
        
        vbcModule.CodeModule.AddFromFile strFile$
        ' import code from the given file

        DoCmd.Save acModule, strModule$
        ' save the module
        
        Module_Overwrite = True
        ' return success to origin function
    
    Else
    ' if the given file does not exist, ...
    
        Module_Overwrite = False
        ' return failure to origin function
        
    End If
    
Else
' if the module was not found, ...

    Module_Overwrite = False
    ' return failure to origin function
    
End If

End Function

Public Function Module_AcccurateCountOfLines(vbcModule As VBComponent) As Long
' provide an accurate count of lines for a given module,
' ignoring comments and blank lines

Dim lngLine As Long
' variable for line being looked at

If vbcModule.Collection.Parent.Protection = vbext_pp_locked Then Module_AcccurateCountOfLines = -1: Exit Function
' if the parent of the module is locked, exit to avoid an error

For lngLine& = 1 To vbcModule.CodeModule.CountOfLines
' from the first to the last of the lines in the code module, ...

    If (Not Trim(vbcModule.CodeModule.Lines(lngLine&, 1)) = vbNullString) And _
        (Not Left(Trim(vbcModule.CodeModule.Lines(lngLine&, 1)), 1) = "'") Then
    ' if not a blank or comment line, ...
    
        Module_AcccurateCountOfLines = Module_AcccurateCountOfLines + 1
        ' increment the count of lines
    
    End If

Next lngLine&
' continue to the next line in the code module

End Function

Public Function IsModule(strModuleName As String) As Variant
' verify if a module exists by name, and provide module as component if so

Dim lngProjectIndex As Long
' variable for project index in projects

Dim vbcComponent As VBComponent
' variable for component in project

Dim varResult(2)
' result of attempt to find module by name
'
' 0 = true/false
' 1 = project
' 2 = module

For lngProjectIndex& = 1 To Application.VBE.VBProjects.Count
' for each project in the vb projects, ...

    For Each vbcComponent In Application.VBE.VBProjects(lngProjectIndex&).VBComponents
    ' for each component in the project, ...

        If LCase(vbcComponent.Name) = LCase(strModuleName$) Then
        ' if the component name matches the module name provided, ...
        
            varResult(0) = True
            ' capture success
            
            Set varResult(1) = Application.VBE.VBProjects(lngProjectIndex&)
            ' capture project
            
            Set varResult(2) = vbcComponent
            ' capture component
            
            IsModule = varResult()
            ' return success, project, and component to origin function
            
        End If
    
    Next vbcComponent
    ' continue to the next component in the project
    
Next lngProjectIndex&
' continue to the next project

If (varResult(0) <> True) Then
' if the module name was not found (and was not explicitly set true), ...

    varResult(0) = False
    ' capture failure
    
    IsModule = varResult()
    ' return failure to origin function

End If

End Function

