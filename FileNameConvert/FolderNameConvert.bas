Attribute VB_Name = "FolderNameConvert"
Option Explicit

Public Sub GetFolderName()

Dim Fso As Object
Dim FolderPath As String
Dim ShortPath As String
Dim TargetFolder As Object
Dim TargetSubfolder As Object
Dim TargetRow As Long

    With ThisWorkbook.Sheets("フォルダ名取得")
  
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
            
            'フォルダ内のサブフォルダを処理
            For Each TargetSubfolder In TargetFolder.SubFolders
            
                .Cells(TargetRow, 1).Value = FolderPath & "\" & TargetSubfolder.Name
                .Cells(TargetRow, 2).Value = TargetSubfolder.Name
                
                TargetRow = TargetRow + 1
                
            Next TargetSubfolder
        
        Else
        
            MsgBox FolderPath & "は存在しません。"
        
        End If
        
    End With
    
    Set TargetSubfolder = Nothing
    Set TargetFolder = Nothing
    Set Fso = Nothing
    
    
End Sub

Public Sub ChangeFolderName()

Dim Fso As Object
Dim FolderFullPath As String
Dim ChangePath As String
Dim newFolderName As String
Dim TargetRow As Long

    TargetRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
             
    With ThisWorkbook.Sheets("フォルダ名変更")
  
        Do
        
            '変更前のフォルダのパスを指定
            FolderFullPath = .Cells(TargetRow, "A").Value
            
            '変更後のフォルダ名
            newFolderName = .Cells(TargetRow, "B").Value
            
            ChangePath = ChangeShortPath(FolderFullPath)
            
            If ChangePath <> "" Then
                
                'フォルダ名を変更
                On Error Resume Next
                
                Fso.GetFolder(ChangePath).Name = newFolderName
            
                If Err.Number <> 0 Then
                    
                    MsgBox FolderFullPath & "の変更に失敗しました。"
                    Err.Clear
                
                End If
                
                On Error GoTo 0

            Else
                
                MsgBox FolderFullPath & "は存在しません。"
            
            End If
            
            TargetRow = TargetRow + 1
             
         Loop Until .Cells(TargetRow, "A").Value = ""
    
    End With
    
    Set Fso = Nothing

    MsgBox "完了しました"
    
End Sub

