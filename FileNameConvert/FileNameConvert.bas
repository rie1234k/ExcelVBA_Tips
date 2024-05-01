Attribute VB_Name = "FileNameConvert"
Option Explicit

'方法① FileSystemObject 長いパスに対応
Public Sub ファイル名取得()

Dim FolderName As String
Dim endRow As Long
Dim Fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long
Dim i As Long
Dim f As Object


    '取得場所
    FolderName = Range("G1").Value
        
    'データ最終行
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'データ消去
    Range(Cells(2, "A"), Cells(endRow, "C")).ClearContents
    
    
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
    
    
        Cells(iRow, 1).Value = FolderName & "\" & iFile.Name
        Cells(iRow, 2).Value = FolderName
        Cells(iRow, 3).Value = iFile.Name
        
        iRow = iRow + 1
        
    Next iFile
    
    
    
    
End Sub

'方法② コマンドライン
Public Sub cmd_ファイル名取得()
Dim wsh As New IWshRuntimeLibrary.WshShell

Dim result As WshExec

Dim cmd As String
Dim filedata() As String
Dim filenm As Variant

Dim endRow As Long

Dim i As Long

endRow = Range("A1").End(xlDown).Row

Range(Range("A2"), Cells(endRow, 3)).ClearContents



cmd = "dir " & """" & Range("G1").Value & """" & " /A:-D /B"

Set result = wsh.Exec("%ComSpec% /c " & cmd)

Do While result.Status = 0
    DoEvents
Loop

filedata = Split(result.StdOut.ReadAll, vbCrLf)

i = 2

For Each filenm In filedata
    
    If filenm <> "" Then Cells(i, 1).Value = Range("G1").Value & "\" & filenm
    If filenm <> "" Then Cells(i, 2).Value = Range("G1").Value
    
    Cells(i, 3).Value = filenm
    
    i = i + 1
Next

Set result = Nothing

Set wsh = Nothing


End Sub

'方法③ Dir
Public Sub dir_ファイル名取得()

Dim FolderName As String
Dim endRow As Long
Dim FileName As String

Dim i As Long

   '取得場所
    FolderName = Range("G1").Value
        
    'データ最終行
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'データ消去
    Range(Cells(2, "A"), Cells(endRow, "C")).ClearContents
  
    FileName = Dir(FolderName & "\*.*")
    
    i = 2
    
    Do While FileName <> ""
    
        Cells(i, 3).Value = FileName
        Cells(i, 2).Value = FolderName
        Cells(i, 1).Value = FolderName & "\" & FileName
        
        i = i + 1
        FileName = Dir()
        
    
    Loop
    
  
End Sub

Public Sub ファイル名変更()  '長いパス対応
Dim Fso As Object

Dim fileFullPath As String
Dim newfileName As String
Dim FolderName As String
Dim FolderShortName As String
Dim iRow As Long

    iRow = 2
    
    With ThisWorkbook.Sheets("ファイル名変更")
  
    
        Do
        
             '変更前のフルパスを指定
             fileFullPath = .Cells(iRow, "A").Value
             
             '変更後のファイル名
             newfileName = .Cells(iRow, "B").Value
             
             
             Set Fso = CreateObject("Scripting.FileSystemObject")
             
             If Len(fileFullPath) > 259 Then
                
            
                fileFullPath = ChangeShortPath(fileFullPath)
                             
             End If
             
             'ファイル名を変更
             Fso.GetFile(fileFullPath).Name = newfileName
             
             '後片付け
             Set Fso = Nothing
             
         
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     MsgBox "完了しました"
     
End Sub
