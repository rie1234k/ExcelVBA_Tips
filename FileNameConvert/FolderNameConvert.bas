Attribute VB_Name = "FolderNameConvert"
Option Explicit

'方法① FileSystemObject　長いパスに対応
Public Sub fso_GetFolderName()
Dim FolderName As String
Dim endRow As Long
Dim Fso As Object
Dim iFolder As Object
Dim subFolder As Object
Dim iRow As Long
Dim i As Long
Dim f As Object


    With ThisWorkbook.Sheets("フォルダ名取得")
  
        '取得場所
        FolderName = .Range("G1").Value
            
        '最終行
        endRow = .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
        
        'データ消去
        .Range(.Cells(2, "A"), .Cells(endRow, "B")).ClearContents
        
        
        Set Fso = CreateObject("Scripting.FileSystemObject")
        
        'フォルダの取得
        Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
        
        For Each f In iFolder.SubFolders: i = i + 1: Next f
        
        If iFolder.SubFolders.Count <> i Then
        
            Set iFolder = Fso.GetFolder(iFolder.ShortPath)
        
        End If
        
        
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


Public Sub ChangeFolderName()  '長いパス対応

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
    
           
           FolderFullPath = ChangeShortPath(FolderFullPath)
            
            'フォルダ名を変更
            Fso.GetFolder(FolderFullPath).Name = newFolderName
        
            iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
    
    
    MsgBox "完了しました"
    
    Set Fso = Nothing
    
End Sub

