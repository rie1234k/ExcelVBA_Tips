Attribute VB_Name = "D_ファイルダウンロード"
Option Explicit

'ファイルダウンロード API宣言
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'キャッシュ削除 API宣言
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" _
(ByVal lpszUrlName As String) As Long


Public Function GetDownloadFilePath(TragetURL As String) As String

Dim iFlag As Long

Dim SaveFilePath As String
Dim SaveFileName As String
  
    '------- ネットワーク上からファイルをダウンロード -------
    SaveFileName = Right(TragetURL, Len(TragetURL) - InStrRev(TragetURL, "/"))
    
    SaveFilePath = ThisWorkbook.Path & "\" & SaveFileName
    
    Call DeleteUrlCacheEntry(TragetURL) 'キャッシュクリア
    
    iFlag = URLDownloadToFile(0, TragetURL, SaveFilePath, 0, 0)
    
    If iFlag <> 0 Then MsgBox "ダウンロード失敗": End
    
    GetDownloadFilePath = SaveFilePath
    
End Function




