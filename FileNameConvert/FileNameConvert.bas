Attribute VB_Name = "FileNameConvert"
Option Explicit

'方法① FileSystemObject 長いパスに対応
Public Sub fso_GetFileName()

Dim FolderName As String
Dim Fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long

Dim i As Long
Dim f As Object

     
    With ThisWorkbook.Sheets("ファイル名取得")
 
        '取得場所
        FolderName = .Range("G1").Value
              
        'データ消去
        .Range("A1").CurrentRegion.Offset(1).ClearContents

        
        Set Fso = CreateObject("Scripting.FileSystemObject")
        
        'フォルダの取得
        Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
            
        
        For Each f In iFolder.Files: i = i + 1: Next f
        
        If iFolder.Files.Count <> i Then
        
            Set iFolder = Fso.GetFolder(iFolder.ShortPath)
        
        End If
        
       
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


Public Sub ChangeFileName()  '長いパス対応
Dim Fso As Object

Dim FileFullPath As String
Dim newfileName As String
Dim FolderName As String
Dim FolderShortName As String
Dim iRow As Long

    iRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
     
    With ThisWorkbook.Sheets("ファイル名変更")
      
        Do
        
             '変更前のフルパスを指定
             FileFullPath = .Cells(iRow, "A").Value
             
             '変更後のファイル名
             newfileName = .Cells(iRow, "B").Value

             FileFullPath = ChangeShortPath(FileFullPath)

             'ファイル名を変更
             Fso.GetFile(FileFullPath).Name = newfileName
            
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     Set Fso = Nothing
     
     MsgBox "完了しました"
     
End Sub
