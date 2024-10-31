Attribute VB_Name = "FileNameConvert"
Option Explicit

Public Sub GetFileName()

Dim FolderName As String
Dim Fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long
 
    With ThisWorkbook.Sheets("ファイル名取得")
 
        '取得場所
        FolderName = .Range("G1").Value
              
        'データ消去
        .Range("A1").CurrentRegion.Offset(1).ClearContents
  
        Set Fso = CreateObject("Scripting.FileSystemObject")
        
        'フォルダの取得
        Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
       
        '行数
        iRow = 2
        
        'フォルダ内のファイルを処理
        For Each iFile In iFolder.Files
        
            .Cells(iRow, 1).Value = FolderName & "\" & iFile.Name
            .Cells(iRow, 2).Value = FolderName
            .Cells(iRow, 3).Value = iFile.Name
            
            iRow = iRow + 1
            
        Next iFile
    
    End With
    
    Set Fso = Nothing
        
End Sub

Public Sub ChangeFileName()

Dim Fso As Object
Dim FileFullPath As String
Dim newFileName As String
Dim iRow As Long

    iRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
     
    With ThisWorkbook.Sheets("ファイル名変更")
      
        Do
        
             '変更前のフルパスを指定
             FileFullPath = .Cells(iRow, "A").Value
             
             '変更後のファイル名
             newFileName = .Cells(iRow, "B").Value
             
             'ショートパスに変換
             FileFullPath = ChangeShortPath(FileFullPath)

             'ファイル名を変更
             Fso.GetFile(FileFullPath).Name = newFileName
            
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     Set Fso = Nothing
     
     MsgBox "完了しました"
     
End Sub
