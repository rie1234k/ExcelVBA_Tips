Attribute VB_Name = "Module1"
Option Explicit

'�t�@�C���_�E�����[�h API�錾

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'�L���b�V���폜 API�錾
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" _
(ByVal lpszUrlName As String) As Long



Public Sub FileDownload()

Dim iFlag As Long '�t�@�C���_�E�����[�h�m�F�p

Dim iURL As String '�_�E�����[�hURL

Dim FilePath As String '�t�@�C���p�X
Dim FileName As String '�t�@�C����

Dim Extension As String '�g���q

Dim i As Long

Dim iCount As Long '����


    '��ʍX�V�A�����v�Z���~
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
        
            Call DeleteUrlCacheEntry(iURL) '�L���b�V���N���A
            
            '�t�@�C���_�E�����[�h
            iFlag = URLDownloadToFile(0, iURL, FilePath, 0, 0)
            
            If iFlag = 0 Then
                
                .Cells(i, "C").Value = Format(Date, "yyyy/mm/dd ") & Format(Time, "hh:mm:ss ") & "����"
            
            Else
            
                .Cells(i, "C").Value = Format(Date, "yyyy/mm/dd ") & Format(Time, "hh:mm:ss ") & "���s"
            
            End If
                        
        
            i = i + 1
            
        Loop Until .Cells(i, "A").Value = ""
        
        
    End With
        
    
    '��ʍX�V�A�����X�V��߂�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
 
    
End Sub

