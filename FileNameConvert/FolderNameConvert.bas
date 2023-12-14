Attribute VB_Name = "FolderNameConvert"
Option Explicit

'方法① FileSystemObject
Public Sub フォルダ名取得()
Dim FolderName As String
Dim endRow As Long
Dim fso As Object
Dim iFolder As Object
Dim subFolder As Object
Dim iRow As Long

    '取得場所
    FolderName = Range("G1").Value
        
    '最終行
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'データ消去
    Range(Cells(2, "A"), Cells(endRow, "B")).ClearContents
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'フォルダの取得
    Set iFolder = fso.GetFolder(FolderName)
    
    '行数
    iRow = 2
    
    'フォルダ内のサブフォルダを処理
    For Each subFolder In iFolder.SubFolders
    
    
        Cells(iRow, 1).Value = subFolder.Path
        Cells(iRow, 2).Value = subFolder.Name
        
        iRow = iRow + 1
        
    Next subFolder
    
    
    

End Sub

'方法② コマンドライン
Public Sub cmd_フォルダ名取得()
Dim wsh As New IWshRuntimeLibrary.WshShell

Dim result As WshExec

Dim cmd As String
Dim filedata() As String
Dim filenm As Variant

Dim endRow As Long

Dim i As Long

endRow = Range("A1").End(xlDown).Row

Range(Range("A2"), Cells(endRow, 2)).ClearContents



cmd = "dir " & """" & Range("G1").Value & """" & " /A:D /B"

Set result = wsh.Exec("%ComSpec% /c " & cmd)

Do While result.Status = 0
    DoEvents
Loop

filedata = Split(result.StdOut.ReadAll, vbCrLf)

i = 2

For Each filenm In filedata
    
    If filenm <> "" Then Cells(i, 1).Value = Range("G1").Value & "\" & filenm
    
    Cells(i, 2).Value = filenm
    
    i = i + 1
Next

Set result = Nothing

Set wsh = Nothing


End Sub

Public Sub フォルダ名変更()

Dim fso As Object

Dim folderFullPath As String
Dim newFolderName As String
Dim iRow As Long

    iRow = 2
    
  With ThisWorkbook.Sheets("フォルダ名変更")
  
    
        Do
        
             '変更前のフォルダのパスを指定
             folderFullPath = .Cells(iRow, "A").Value
             
             '変更後のフォルダ名
             newFolderName = .Cells(iRow, "B").Value
             
             
             Set fso = CreateObject("Scripting.FileSystemObject")
             
             'フォルダ名を変更
             fso.GetFolder(folderFullPath).Name = newFolderName
             
             '後片付け
             Set fso = Nothing
             
         
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
    
    MsgBox "完了しました"
    
End Sub

