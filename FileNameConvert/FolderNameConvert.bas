Attribute VB_Name = "FolderNameConvert"
Option Explicit

Public Sub GetFolderName()

Dim FolderName As String
Dim Fso As Object
Dim iFolder As Object
Dim subFolder As Object
Dim iRow As Long

    With ThisWorkbook.Sheets("フォルダ名取得")
  
        '取得場所
        FolderName = .Range("G1").Value
            
        'データ消去
        .Range("A1").CurrentRegion.Offset(1).ClearContents
        
        Set Fso = CreateObject("Scripting.FileSystemObject")
        
        'フォルダの取得
        Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
                
        '行数
        iRow = 2
        
        'フォルダ内のサブフォルダを処理
        For Each subFolder In iFolder.SubFolders
        
            .Cells(iRow, 1).Value = FolderName & "\" & subFolder.Name
            .Cells(iRow, 2).Value = subFolder.Name
            
            iRow = iRow + 1
            
        Next subFolder
        
    End With
    
    Set Fso = Nothing
    
    
End Sub

Public Sub ChangeFolderName()

Dim Fso As Object
Dim FolderFullPath As String
Dim newFolderName As String
Dim iRow As Long

    iRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
             
    With ThisWorkbook.Sheets("フォルダ名変更")
  
        Do
        
            '変更前のフォルダのパスを指定
            FolderFullPath = .Cells(iRow, "A").Value
            
            '変更後のフォルダ名
            newFolderName = .Cells(iRow, "B").Value
    
            'ショートパスに変換
            FolderFullPath = ChangeShortPath(FolderFullPath)
            
            'フォルダ名を変更
            Fso.GetFolder(FolderFullPath).Name = newFolderName
        
            iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
    
    Set Fso = Nothing

    MsgBox "完了しました"
    
End Sub

