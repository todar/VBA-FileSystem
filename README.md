# VBA-FileSystem
FileSystem functions to make it easier and faster coding. This is really handy for users who are maybe new at VBA, or just looking to do simple tasks faster.

# Public Funtions:
- AppendTextFile
- CreateFilePath
- CreateFolder
- DeleteFile
- DeleteFolder
- FileExists
- FolderExists
- HasWriteAccessToFolder
- MoveFile
- OpenAnyFile
- OpenFileExplorer
- OpenURL
- ReadTextFile

# Usage

Import FileSystemFunctions.bas file.

Set Reference to MicroSoft Scripting Runtime (for using Scripting.FileSystemObject)

# Example

```vb
 Private Sub FileSystemFunctionExamples()
    
    HasWriteAccessToFolder ("C:\Program Files") '-> True || False
    
    'The other functions are pretty self explanatory :)
    
End Sub

```
