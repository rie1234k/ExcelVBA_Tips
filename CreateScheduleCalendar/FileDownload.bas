Attribute VB_Name = "D_�t�@�C���_�E�����[�h"
Option Explicit

'�t�@�C���_�E�����[�h API�錾
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'�L���b�V���폜 API�錾
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" _
(ByVal lpszUrlName As String) As Long


Public Function GetDownloadFilePath(TragetURL As String) As String

Dim iFlag As Long

Dim SaveFilePath As String
Dim SaveFileName As String
  
    '------- �l�b�g���[�N�ォ��t�@�C�����_�E�����[�h -------
    SaveFileName = Right(TragetURL, Len(TragetURL) - InStrRev(TragetURL, "/"))
    
    SaveFilePath = ThisWorkbook.Path & "\" & SaveFileName
    
    Call DeleteUrlCacheEntry(TragetURL) '�L���b�V���N���A
    
    iFlag = URLDownloadToFile(0, TragetURL, SaveFilePath, 0, 0)
    
    If iFlag <> 0 Then MsgBox "�_�E�����[�h���s": End
    
    GetDownloadFilePath = SaveFilePath
    
End Function




