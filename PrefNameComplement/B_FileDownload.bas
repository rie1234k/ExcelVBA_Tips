Attribute VB_Name = "B_FileDownload"
Option Explicit

'ファイルダウンロード API宣言
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'キャッシュ削除 API宣言
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" _
(ByVal lpszUrlName As String) As Long


Public Sub PostcodeZipFileDowunload(TragetURL As String, PostcodeFileName As String)

Dim iFlag As Long

Dim SaveFilePath As String
Dim SaveFileName As String
    
Dim psCommand As String
Dim Wsh As Object

  
    '------- ネットワーク上から郵便番号データをダウンロード -------
    SaveFileName = Right(TragetURL, Len(TragetURL) - InStrRev(TragetURL, "/"))
    
    SaveFilePath = ThisWorkbook.Path & "\" & SaveFileName
    
    Call DeleteUrlCacheEntry(TragetURL) 'キャッシュクリア
    
    iFlag = URLDownloadToFile(0, TragetURL, SaveFilePath, 0, 0)
    
    If iFlag <> 0 Then MsgBox "ダウンロード失敗": End
    
   
   '------- Zipファイルを展開 -------
    psCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Path " & SaveFilePath & " -DestinationPath " & ThisWorkbook.Path & " -Force"
    
    Set Wsh = CreateObject("WScript.Shell")

    iFlag = Wsh.Run(Command:=psCommand, WindowStyle:=0, WaitOnReturn:=True)
    
     If iFlag <> 0 Then MsgBox "Zip解凍失敗": End
     
    Set Wsh = Nothing
    
    
End Sub
