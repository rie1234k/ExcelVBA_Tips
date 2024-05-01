Attribute VB_Name = "FolderNameConvert"
Option Explicit

'���@�@ FileSystemObject�@�����p�X�ɑΉ�
Public Sub �t�H���_���擾()
Dim FolderName As String
Dim endRow As Long
Dim Fso As Object
Dim iFolder As Object
Dim subFolder As Object
Dim iRow As Long
Dim i As Long
Dim f As Object

    '�擾�ꏊ
    FolderName = Range("G1").Value
        
    '�ŏI�s
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    '�f�[�^����
    Range(Cells(2, "A"), Cells(endRow, "B")).ClearContents
    
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    '�t�H���_�̎擾
    Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
    
    For Each f In iFolder.SubFolders: i = i + 1: Next f
    
    If iFolder.SubFolders.Count <> i Then
    
        Set iFolder = Fso.GetFolder(iFolder.ShortPath)
    
    End If
    
    
    '�s��
    iRow = 2
    
    '�t�H���_���̃T�u�t�H���_������
    For Each subFolder In iFolder.SubFolders
    
    
        Cells(iRow, 1).Value = FolderName & "\" & subFolder.Name
        Cells(iRow, 2).Value = subFolder.Name
        
        iRow = iRow + 1
        
    Next subFolder
    
    

End Sub

'���@�A �R�}���h���C��
Public Sub cmd_�t�H���_���擾()
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

Public Sub �t�H���_���ύX()  '�����p�X�Ή�

Dim Fso As Object

Dim folderFullPath As String
Dim newFolderName As String
Dim iRow As Long

    iRow = 2
    
    With ThisWorkbook.Sheets("�t�H���_���ύX")
  
    
        Do
        
             '�ύX�O�̃t�H���_�̃p�X���w��
             folderFullPath = .Cells(iRow, "A").Value
             
             '�ύX��̃t�H���_��
             newFolderName = .Cells(iRow, "B").Value
             
             
             Set Fso = CreateObject("Scripting.FileSystemObject")
             
             If Len(folderFullPath) > 259 Then
                
            
                folderFullPath = ChangeShortPath(folderFullPath)
                             
             End If
             
             
             '�t�H���_����ύX
             Fso.GetFolder(folderFullPath).Name = newFolderName
             
             '��Еt��
             Set Fso = Nothing
             
         
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
    
    MsgBox "�������܂���"
    
    
End Sub

