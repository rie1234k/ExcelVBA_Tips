Attribute VB_Name = "FileNameConvert"
Option Explicit

Public Sub GetFileName()

Dim Fso As Object
Dim FolderPath As String
Dim ShortPath As String
Dim TargetFolder As Object
Dim TargetFile As Object
Dim TargetRow As Long
 
    With ThisWorkbook.Sheets("ファイル名取得")
 
        '取得場所
        FolderPath = .Range("G1").Value
              
        'データ消去
        .Range("A1").CurrentRegion.Offset(1).ClearContents
  
        Set Fso = CreateObject("Scripting.FileSystemObject")
         
        ShortPath = ChangeShortPath(FolderPath)
        
        If ShortPath <> "" Then
            
            'フォルダの取得
            Set TargetFolder = Fso.GetFolder(ShortPath)
        
            '行数
            TargetRow = 2
            
            'フォルダ内のファイルを処理
            For Each TargetFile In TargetFolder.Files
            
                .Cells(TargetRow, 1).Value = FolderPath & "\" & TargetFile.Name
                .Cells(TargetRow, 2).Value = FolderPath
                .Cells(TargetRow, 3).Value = TargetFile.Name
                
                TargetRow = TargetRow + 1
                
            Next TargetFile
        
        Else
            
            MsgBox FolderPath & "は存在しません。"
        
        End If
 
    End With
    
    Set TargetFile = Nothing
    Set TargetFolder = Nothing
    Set Fso = Nothing
    
        
End Sub

Public Sub ChangeFileName()

Dim Fso As Object
Dim FileFullPath As String
Dim ChangePath As String
Dim newFileName As String
Dim TargetRow As Long

    TargetRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
     
    With ThisWorkbook.Sheets("ファイル名変更")
      
        Do
        
             '変更前のフルパスを指定
             FileFullPath = .Cells(TargetRow, "A").Value
             
             '変更後のファイル名
             newFileName = .Cells(TargetRow, "B").Value
             
             'ショートパスに変換 ChangeShortPathは、存在しない場合、空白を返す
             ChangePath = ChangeShortPath(FileFullPath)
             
             If ChangePath <> "" Then
                
                
                'ファイル名を変更
                On Error Resume Next
                
                Fso.GetFile(ChangePath).Name = newFileName
                
                If Err.Number <> 0 Then
                    
                    MsgBox FileFullPath & "の変更に失敗しました。"
                    Err.Clear
                
                End If
                
                On Error GoTo 0
                
             Else
             
                MsgBox FileFullPath & "は存在しません。"
             
             End If
             
             TargetRow = TargetRow + 1
             
         Loop Until .Cells(TargetRow, "A").Value = ""
    
    End With
     
    Set Fso = Nothing
     
    MsgBox "完了しました"
     
End Sub
