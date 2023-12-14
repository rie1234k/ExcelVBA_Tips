Attribute VB_Name = "FileNameConvert"
Option Explicit

'方法① FileSystemObject
Public Sub ファイル名取得()

Dim FolderName As String
Dim endRow As Long
Dim fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long

    '取得場所
    FolderName = Range("G1").Value
        
    'データ最終行
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'データ消去
    Range(Cells(2, "A"), Cells(endRow, "C")).ClearContents
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'フォルダの取得
    Set iFolder = fso.GetFolder(FolderName)
    
    '行数
    iRow = 2
    
    'フォルダ内のファイルを処理
    For Each iFile In iFolder.Files
    
    
        Cells(iRow, 1).Value = iFile.Path
        Cells(iRow, 2).Value = iFile.ParentFolder
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

Public Sub ファイル名変更()
Dim fso As Object

Dim fileFullPath As String
Dim newfileName As String
Dim iRow As Long

    iRow = 2
    
    With ThisWorkbook.Sheets("ファイル名変更")
  
    
        Do
        
             '変更前のフルパスを指定
             fileFullPath = .Cells(iRow, "A").Value
             
             '変更後のファイル名
             newfileName = .Cells(iRow, "B").Value
             
             
             Set fso = CreateObject("Scripting.FileSystemObject")
             
             'ファイル名を変更
             fso.GetFile(fileFullPath).Name = newfileName
             
             '後片付け
             Set fso = Nothing
             
         
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     MsgBox "完了しました"
     
End Sub

