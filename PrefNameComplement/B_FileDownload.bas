Attribute VB_Name = "B_FileDownload"
Option Explicit

'ファイルダウンロード API宣言
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long

'キャッシュ削除 API宣言
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" _
(ByVal lpszUrlName As String) As Long

Public Enum AppError
    
    BaseNo = vbObjectError + 1000
    
    DownloadFailed = BaseNo + 1
    FileNotFound = BaseNo + 2
    ZipFailed = BaseNo + 3
    
End Enum


Public Sub PostcodeZipFileDownload(TargetURL As String)

Dim iFlag As Long
Dim SaveFilePath As String
Dim SaveFileName As String
Dim psCommand As String
Dim Wsh As Object

Dim errNum As Long
Dim errSrc As String
Dim errDesc As String

Const CodeName As String = "PostcodeZipFileDownload"

    On Error GoTo ErrHandler
  
    '------- ネットワーク上から郵便番号データをダウンロード -------
    SaveFileName = Mid(TargetURL, InStrRev(TargetURL, "/") + 1)
    
    SaveFilePath = ThisWorkbook.Path & "\" & SaveFileName
    
    DeleteUrlCacheEntry TargetURL 'キャッシュクリア
    
    iFlag = URLDownloadToFile(0, TargetURL, SaveFilePath, 0, 0)
    
    If iFlag <> 0 Then Err.Raise AppError.DownloadFailed, CodeName, "ファイルダウンロード失敗" & vbCrLf & "(" & TargetURL & ")"
    
    'ダウンロードファイルの存在を確認
    If Dir(SaveFilePath) = "" Then Err.Raise AppError.FileNotFound, CodeName, "ダウンロードファイル保存失敗" & vbCrLf & "(" & SaveFilePath & ")"
    
   '------- Zipファイルを展開 -------
    psCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Path '" & SaveFilePath & "' -DestinationPath '" & ThisWorkbook.Path & "' -Force"
    
    Set Wsh = CreateObject("WScript.Shell")

    iFlag = Wsh.Run(Command:=psCommand, WindowStyle:=0, WaitOnReturn:=True)
    
    If iFlag <> 0 Then Err.Raise AppError.ZipFailed, CodeName, "ZIPファイル解凍失敗" & vbCrLf & "(" & SaveFilePath & ")"
     
     GoTo Cleanup

ErrHandler:
    
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description

    Resume Cleanup

Cleanup:

    On Error GoTo 0
    
    Set Wsh = Nothing
    
    If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc

    
End Sub
