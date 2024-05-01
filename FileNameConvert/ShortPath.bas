Attribute VB_Name = "ShortPath"
Option Explicit

Public Function ChangeShortPath(FullPath As String) As String

Dim Fso As Object
Dim TargetFolder As String
Dim endPath As String
Dim FolderPath As String

    endPath = ""
    TargetFolder = ""
    FolderPath = ""

    Set Fso = CreateObject("Scripting.FileSystemObject")
        
        
    TargetFolder = FullPath
    
    '存在するフォルダまで遡って、存在するフォルダをショートパスに変換する
    Do Until Fso.FolderExists(TargetFolder)
        
        If endPath <> "" Then
        
            endPath = Fso.GetFileName(TargetFolder) & "\" & endPath
        
        Else
        
            endPath = Fso.GetFileName(TargetFolder)
        
        End If
        
        TargetFolder = Fso.GetParentFolderName(TargetFolder)
        
    
    Loop
    
    
    If endPath <> "" Then
        
        TargetFolder = Fso.GetFolder(TargetFolder).ShortPath & "\" & endPath
         
    End If
    
    
    '変換したショートパスと末尾のパスを繋げたパスについて、長い場合は再度ショートパスに変換する
    endPath = ""
    
    Do Until Fso.FolderExists(TargetFolder)
        
        If endPath <> "" Then
        
            endPath = Fso.GetFileName(TargetFolder) & "\" & endPath
        
        Else
        
            endPath = Fso.GetFileName(TargetFolder)
        
        End If
        
        TargetFolder = Fso.GetParentFolderName(TargetFolder)
        
    
    Loop
    
    
    If endPath <> "" Then
        
        ChangeShortPath = Fso.GetFolder(TargetFolder).ShortPath & "\" & endPath
    
    Else
    
         ChangeShortPath = TargetFolder
         
    End If
    

End Function


