Attribute VB_Name = "ShortPath"
Option Explicit

Public Function ChangeShortPath(FullPath As String) As String

Dim Fso As Object
Dim TargetPath As String
Dim LastPath As String

Dim i As Long

   Set Fso = CreateObject("Scripting.FileSystemObject")
        
    TargetPath = FullPath
    
    '存在するフォルダまで遡って、存在するフォルダをショートパスに変換する
    Do Until Fso.FolderExists(TargetPath)
    
        'GetFileNameは、ファイルに限らず、最終要素と取り出す
        LastPath = "\" & Fso.GetFileName(TargetPath) & LastPath
        
        TargetPath = Fso.GetParentFolderName(TargetPath)
        
        If TargetPath = "" Then
            
            MsgBox "指定されたフォルダは存在しません。"
            End
            
        End If
        
    Loop
    
    TargetPath = Fso.GetFolder(TargetPath).ShortPath & LastPath
    
    ChangeShortPath = TargetPath
        
    If Fso.FileExists(ChangeShortPath) = False And Fso.FolderExists(ChangeShortPath) = False Then
        
        MsgBox "指定されたパスは存在しません。"
        End
        
    End If
    
    Set Fso = Nothing
    
End Function
