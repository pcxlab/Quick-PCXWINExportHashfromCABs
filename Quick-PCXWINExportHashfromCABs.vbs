' ==============================================================================
' Synopsis: This VBScript code extracts hash files from CAB archives found in the
'           current directory. If a hash file with the same name already exists 
'           in the destination folder, it renames the new file with a unique 
'           suffix before copying it. Once extracted, the hash files are moved to
'           the same directory as the script. This script requires the presence
'           of CAB archives containing hash files in the current directory for 
'           successful execution. 
'
' Title: Extracting Hash Files from CAB Archives
'
' Author/Creator/Researcher: Roopesh C Shet (roopeshcshet88@gmail.com)
'
' Introduction/Background: Extracting hash files from CAB archives can be a time-consuming process, often requiring manual effort.
'
' Main Plot/Thesis: This tool aims to streamline the process of extracting hash files from CAB archives.
'
' Key Points/Arguments:
'   - Significantly reduces time and effort required for extracting hash files.
'   - Minimizes the likelihood of errors during the extraction process.
'
' Conclusion/Summary: The tool provides a simple solution, enabling hash file extraction in just two clicks.
'
' Solution: The tool automatically extracts hash files from CAB archives, simplifying the process.
'
' Benefit: Users will save time and effort while ensuring error-free execution of hash file extraction tasks.
' 
' Note: Ensure that this script is placed in the same directory as the CAB 
'       archives containing hash files. The script will automatically extract 
'       and rename hash files from all CAB archives found in the current 
'       directory.
'
' Date: 12-May-2024
' ==============================================================================

' ==============================================================================
' Synopsis: Extracts hash files from CAB archives in the current directory.
'           Renames hash files if duplicates exist, then moves them to the script's directory.
'           Requires CAB archives in the current directory for execution.
'
' Title: Extracting Hash Files from CAB Archives
' Author: Roopesh C Shet (roopeshcshet88@gmail.com)
' Date: 09-May-2024
' ==============================================================================

Dim objShell, objFSO, strCurrentDirectory, objFolder, objFile, strFileName
Dim strDestinationFolder, strExternalFolder

Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentDirectory = objFSO.GetAbsolutePathName(".")
Set objFolder = objFSO.GetFolder(strCurrentDirectory)

For Each objFile In objFolder.Files
    If objFSO.GetExtensionName(objFile.Path) = "cab" Then
        strFileName = objFSO.GetBaseName(objFile.Path)
        strDestinationFolder = strCurrentDirectory & "\" & strFileName
        If Not objFSO.FolderExists(strDestinationFolder) Then objFSO.CreateFolder(strDestinationFolder)
        
        objShell.Namespace(strDestinationFolder).CopyHere objShell.Namespace(objFile.Path).Items
        WScript.Sleep 5000
        
        Dim extractedFolder, file, newFileName, count
        Set extractedFolder = objFSO.GetFolder(strDestinationFolder)
        For Each file In extractedFolder.Files
            If objFSO.GetExtensionName(file.Path) = "csv" Then
                strExternalFolder = strCurrentDirectory
                newFileName = objFSO.GetFileName(file.Path)
                count = 1
                Do While objFSO.FileExists(strExternalFolder & "\" & newFileName)
                    newFileName = objFSO.GetBaseName(file.Path) & "_" & count & "." & objFSO.GetExtensionName(file.Path)
                    count = count + 1
                Loop
                objFSO.CopyFile file.Path, strExternalFolder & "\" & newFileName
            End If
        Next
        objFSO.DeleteFolder strDestinationFolder
        objFSO.DeleteFile objFile.Path
    End If
Next

Set objShell = Nothing
Set objFSO = Nothing

