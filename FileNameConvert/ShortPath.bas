Attribute VB_Name = "ShortPath"
Option Explicit

'存在しないパスを渡した場合、空欄""を返す
Public Function ChangeShortPath(FullPath As String) As String

Dim Fso As Object
Dim TargetPath As String
Dim LastPath As String

   Set Fso = CreateObject("Scripting.FileSystemObject")
        
    TargetPath = FullPath
    
    '存在するフォルダまで遡って、存在するフォルダをショートパスに変換する
    Do Until Fso.FolderExists(TargetPath)
    
        'GetFileNameは、ファイルに限らず、最終要素を取り出す
        LastPath = "\" & Fso.GetFileName(TargetPath) & LastPath
        
        'GetParentFolderNameは最終要素のひとつ前の要素を取り出す
        TargetPath = Fso.GetParentFolderName(TargetPath)
        
        'TargetPathが空欄 ＝ フォルダを最初まで遡っても存在しなかった ⇒ 存在しないパスであるため、終了
        If TargetPath = "" Then Exit Do
        
    Loop
    
    TargetPath = Fso.GetFolder(TargetPath).ShortPath & LastPath
    
    If Fso.FileExists(TargetPath) = False _
        And Fso.FolderExists(TargetPath) = False Then TargetPath = ""
    
    ChangeShortPath = TargetPath
    
    Set Fso = Nothing
    
End Function
