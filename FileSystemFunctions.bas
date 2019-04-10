Attribute VB_Name = "FileSystemFunctions"
Option Explicit
Option Compare Text
Option Private Module

'@Author: Robert Todar <robert@roberttodar.com>
'@Licence: MIT

'DEPENDENCIES
' - REFERENCE TO SCRIPTING RUNTIME FOR Scripting.FileSystemObject

'PUBLIC FUNCTIONS
' - AppendTextFile
' - CreateFilePath
' - CreateFolder
' - DeleteFile
' - DeleteFolder
' - FileExists
' - FolderExists
' - MoveFile
' - OpenAnyFile
' - OpenFileExplorer
' - OpenURL
' - ReadTextFile

'NOTES:
'TODO:

'******************************************************************************************
' PUBLIC FUNCTIONS
'******************************************************************************************

Private Sub FileSystemFunctionExamples()
    
    HasWriteAccessToFolder ("C:\Program Files") '-> True || False
    
    'The other functions are pretty self explanatory
    
End Sub
    

'CHECK TO SEE IF CURRENT USER HAS WRITE ACCESS TO FOLDER
Public Function HasWriteAccessToFolder(ByVal FolderPath As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATES A TEMP FILE TO SEE IF USER CAN WRITE TO IT.
    'EXAMPLE: HasWriteAccessToFolder("C:\Program Files") -> True || False
    
    'MAKE SURE FOLDER EXISTS, THIS FUNCTION RETURNS FALSE IF IT DOES NOT
    Dim Fso As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    If Not Fso.FolderExists(FolderPath) Then
        Exit Function
    End If

    'GET UNIQUE TEMP FilePath, DON'T WANT TO OVERWRITE SOMETHING THAT ALREADY EXISTS
    Do
        Dim Count As Integer
        Dim FilePath As String
        
        FilePath = Fso.BuildPath(FolderPath, "TestWriteAccess" & Count & ".tmp")
        Count = Count + 1
    Loop Until Not Fso.FileExists(FilePath)
    
    'ATTEMPT TO CREATE THE TMP FILE, ERROR RETURNS FALSE
    On Error GoTo Catch
    Fso.CreateTextFile(FilePath).Write ("Test Folder Access")
    Kill FilePath
    
    'NO ERROR, ABLE TO WRITE TO FILE; RETURN TRUE!
    HasWriteAccessToFolder = True
    
Catch:
    
End Function

'CALL TO CREATE FILEPATH AND WRITE TO TEXT FILE
Public Sub WriteToTextFile(ByVal FilePath As String, ByVal Value As String)
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New FileSystemObject
    
    If Not Fso.FileExists(FilePath) Then
        CreateFilePath FilePath
    End If
    
    Dim Ts As TextStream
    Set Ts = Fso.OpenTextFile(FilePath, ForWriting, True)
    
    Ts.Write Value
    
End Sub

'READ ANY TEXT FILE, EMPTY STRING IF FILE DOES NOT EXIST
Public Function ReadTextFile(ByVal FilePath As String) As String
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    On Error GoTo NoFile
    Dim Ts As TextStream
    Set Ts = Fso.OpenTextFile(FilePath, ForReading, False)
    
    ReadTextFile = Ts.ReadAll
    
    Exit Function
NoFile:
    'FOR MY NEEDS, ERROR JUST RETURNS AND EMPTY STRING. ADJUST AS NEED FOR OTHER CODE.
    
End Function

'CALL TO CREATE FILEPATH AND APPEND TO TEXT FILE
Public Sub AppendTextFile(ByVal FilePath As String, ByVal Value As String)
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    If Not Fso.FileExists(FilePath) Then
        CreateFilePath FilePath
    End If
    
    Dim Ts As TextStream
    Set Ts = Fso.OpenTextFile(FilePath, ForAppending, True)
    
    Ts.WriteLine Value
    
End Sub

'CREATES FULL PATH. NORMAL CREATE FOLDER OR FILE ONLY DOES ONE LEVEL.
Public Function CreateFilePath(ByVal FullPath As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    'DESC: FSO.CREATEFOLDER && CREATEFILE ONLY CREATE ONE LEVEL, THIS STEPS THROUGH FULL PATH
    
    Dim Paths() As String
    Paths = Split(FullPath, "\")
    
    Dim PathIndex As Integer
    For PathIndex = LBound(Paths, 1) To UBound(Paths, 1) - 1
        
        Dim CurrentPath As String
        CurrentPath = CurrentPath & Paths(PathIndex) & "\"
        
        Dim Fso As New FileSystemObject
        If Not Fso.FolderExists(CurrentPath) Then
            Fso.CreateFolder CurrentPath
        End If
        
    Next PathIndex

End Function

'EASY WAY TO SEE IF FILE EXISTS
Public Function FileExists(ByVal FileSpec As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    FileExists = Fso.FileExists(FileSpec)

End Function

'EASY WAY TO SEE IF FOLDER EXISTS
Public Function FolderExists(ByVal FileSpec As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    FolderExists = Fso.FolderExists(FileSpec)

End Function

'EASY WAY TO CREATE A FOLDER
Public Function CreateFolder(ByVal FolderPath As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    Fso.CreateFolder FolderPath
    
End Function

'EASY WAY TO DELETE A FOLDER
Public Function DeleteFolder(ByVal FolderPath As String)
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    If Fso.FolderExists(FolderPath) Then
        Fso.DeleteFolder FolderPath, True
    End If

End Function

'EASY WAY TO DELETE A FILE
Public Function DeleteFile(ByVal FilePath As String)
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    If Fso.FolderExists(FilePath) Then
        Fso.DeleteFile FilePath, True
    End If

End Function

'EASY WAY TO MOVE A FILE
Public Function MoveFile(ByVal Source As String, ByVal Destination As String)
    
    'AUTHOR: ROBERT TODAR
    'DESC: CREATED FOR EASE OF USE AS THIS CODE IS REPEATED A LOT.
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    Fso.MoveFile Source, Destination
    
End Function

'CHECKS TO SEE IF FILE EXISTS, THEN OPENS IT IF IT DOES
Public Function OpenAnyFile(ByVal FileToOpen As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    'DESC: ABLE TO OPEN ANY FILE OTHER THAN URL.
    
    'WILL ONLY OPEN FILE IF IT EXISTS
    If FileExists(FileToOpen) Then
        OpenAnyFile = True
        
        'API FUNCTION FOR OPENING FILES
        Call ShellExecute(0, "Open", FileToOpen & vbNullString, _
        vbNullString, vbNullString, 1)
    End If
    
End Function

'CHECKS TO SEE IF FOLDER EXISTS, THEN OPENS WINDOWS EXPLORER TO THAT PATH
Public Function OpenFileExplorer(ByVal FolderPath As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    
    If FolderExists(FolderPath) Then
        OpenFileExplorer = True
        Call Shell("explorer.exe " & Chr(34) & FolderPath & Chr(34), vbNormalFocus)
    End If
    
End Function

'OPEN URL IN DEFAULT BROWSER
Public Function OpenURL(ByVal UrlToOpen As String) As Boolean
    
    'AUTHOR: ROBERT TODAR
    
    'API FUNCTION FOR OPENING FILES
    Call ShellExecute(0, "Open", UrlToOpen & vbNullString, _
    vbNullString, vbNullString, 1)
    
End Function
