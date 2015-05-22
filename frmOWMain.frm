Option Explicit

Private Sub btnChooseMap_Click()
    'SetFolderPath txtSavePath

    ' Choose a folder path to output the results
    Dim fd As Office.FileDialog
    Dim strPath As String
    
    ' Turn on Error handling for a potential file problem
    On Error GoTo FileProblem
    

    
    ' Use a folder picker to choose just the folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Set the default directory to be the Minecraft Saves location
    fd.InitialFileName = "%appdata%\.minecraft\saves"
    
    With fd
        If .Show Then
            ' Store the path of the selected folder
            strPath = .SelectedItems(1)
            
            ' Place the path name in the dialog box
            Me.txtSavePath.Value = strPath
        End If
    End With
    
    ' Set error handling by turning it off
    On Error GoTo 0
    
        
    ' Exit the subroutine without executing the error code below
CleanExit:
    Set fd = Nothing
    Exit Sub
    
FileProblem:
    MsgBox "A problem has occurred trying to find the maps and I cannot continue. Please try again.", vbCritical, "Possible File Problem"
    Resume CleanExit
End Sub

Private Sub btnCreateMap_Click()
    ' Run Overviewer using the Shell command and the specified arguments
    Dim strCommand As String
    Dim retVal As Variant
    
    ' strCommand = txtOVPath & "\" & "overviewer " & """%appdata%\.minecraft\saves\" & cboMaps.Value & """" & " " & txtOutputPath.Value
    strCommand = txtOVPath & "\" & "overviewer " & """" & txtSavePath.Value & """" & " " & """" & txtOutputPath.Value & """"
    
    Debug.Print strCommand
    On Error GoTo CantRun
        ' Run Overviewer using the shell command
        retVal = Shell(strCommand, vbNormalFocus)
    On Error GoTo 0
    
    MsgBox "Done"
    
    ' Close the main form
    Unload Me
    Exit Sub
    
CantRun:
    MsgBox "Can't execute the required file.", vbCritical, "File Error"
End Sub

Private Sub btnOutput_Click()
    ' Choose a folder path to output the results
    Dim fd As Office.FileDialog
    Dim strPath As String
    
    ' Turn on Error handling for a potential file problem
    On Error GoTo FileProblem
    
    ' Use a folder picker to choose just the folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show Then
            ' Store the path of the selected folder
            strPath = .SelectedItems(1)
            
            ' Place the path name in the dialog box
            Me.txtOutputPath.Value = strPath
        End If
    End With
    
    ' Set error handling by turning it off
    On Error GoTo 0
    
    ' Exit the subroutine without executing the error code below
CleanExit:
    Set fd = Nothing
    Exit Sub
    
FileProblem:
    MsgBox "A problem has occurred and I cannot continue. Please try again.", vbCritical, "Possible File Problem"
    Resume CleanExit
End Sub

Private Sub SetFolderPath(txtBox As TextBox)
    ' Choose a folder path to output the results
    Dim fd As Office.FileDialog
    Dim strPath As String
    
    ' Turn on Error handling for a potential file problem
    On Error GoTo FileProblem
    
    ' Use a folder picker to choose just the folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show Then
            ' Store the path of the selected folder
            strPath = .SelectedItems(1)
            
            ' Place the path name in the dialog box
            Me.txtBox.Value = strPath
        End If
    End With
    
    ' Set error handling by turning it off
    On Error GoTo 0
    
    ' Exit the subroutine without executing the error code below
CleanExit:
    Set fd = Nothing
    Exit Sub
    
FileProblem:
    MsgBox "A problem has occurred and I cannot continue. Please try again.", vbCritical, "Possible File Problem"
    Resume CleanExit
End Sub

Private Sub btnOVBrowse_Click()
    ' Allow the user to browse to the location where Overviewer is installed.
    Dim fd As Office.FileDialog
    Dim strPath As String
    
    ' Turn on Error handling for a potential file problem
    On Error GoTo FileProblem
    
    ' Use a folder picker to choose just the folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show Then
            ' Store the path of the selected folder
            strPath = .SelectedItems(1)
            
            ' Place the path name in the dialog box
            Me.txtOVPath.Value = strPath
        End If
    End With
    
    ' Set error handling by turning it off
    On Error GoTo 0
    
    ' Exit the subroutine without executing the error code below
    Exit Sub
    
FileProblem:
    MsgBox "A problem has occurred and I cannot continue. Please try again.", vbCritical, "Possible File Problem"
End Sub

Private Sub UserForm_Initialize()
    Dim strDirectory As String
    
    ' Set up some values
    txtOVPath.Value = "C:\Users\Rick\Documents\MinecraftOverviewer\"
    
    On Error GoTo DirProblem
    ' Get the names of the installed worlds from the default Minecraft directory
    'strDirectory = Dir("%appdata%\.minecraft\saves\", vbNormal)
    'Debug.Print "Init: strDirectory = " & strDirectory
    
    'cboMaps.AddItem strDirectory
    cboMaps.AddItem "First World"
    
    txtOutputPath.Value = "C:\maps"

    Exit Sub
    
DirProblem:
    MsgBox "A problem occurred while trying to get the list of worlds."
End Sub
