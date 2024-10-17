Attribute VB_Name = "Module1"
Option Explicit

'ファイルダウンロード API宣言

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'キャッシュ削除 API宣言
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" _
(ByVal lpszUrlName As String) As Long



Public Sub FileDownload()

Dim iFlag As Long 'ファイルダウンロード確認用

Dim iURL As String 'ダウンロードURL

Dim FilePath As String 'ファイルパス
Dim FileName As String 'ファイル名

Dim Extension As String '拡張子

Dim i As Long

Dim iCount As Long '件数


    '画面更新、自動計算を停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        
        
    With ThisWorkbook.Sheets("FileDownload")
    
        Intersect(.Columns("C:C"), .Range("A1").CurrentRegion.Offset(1, 2)).ClearContents
        
        iCount = .Cells(Rows.Count, "A").End(xlUp).Row - 1
    
        i = 2
        
        Do
            iURL = .Cells(i, "A").Value
            
            Extension = Right(iURL, Len(iURL) - InStrRev(iURL, "."))
            
            FileName = .Cells(i, "B").Value & "." & Extension

            FilePath = ThisWorkbook.Path & "\" & FileName
        
            Call DeleteUrlCacheEntry(iURL) 'キャッシュクリア
            
            'ファイルダウンロード
            iFlag = URLDownloadToFile(0, iURL, FilePath, 0, 0)
            
            If iFlag = 0 Then
                
                .Cells(i, "C").Value = Format(Date, "yyyy/mm/dd ") & Format(Time, "hh:mm:ss ") & "成功"
            
            Else
            
                .Cells(i, "C").Value = Format(Date, "yyyy/mm/dd ") & Format(Time, "hh:mm:ss ") & "失敗"
            
            End If
                        
        
            i = i + 1
            
        Loop Until .Cells(i, "A").Value = ""
        
        
    End With
        
    
    '画面更新、自動更新を戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
 
    
End Sub

