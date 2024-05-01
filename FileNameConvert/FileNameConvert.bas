Attribute VB_Name = "FileNameConvert"
Option Explicit

'���@�@ FileSystemObject �����p�X�ɑΉ�
Public Sub �t�@�C�����擾()

Dim FolderName As String
Dim endRow As Long
Dim Fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long
Dim i As Long
Dim f As Object


    '�擾�ꏊ
    FolderName = Range("G1").Value
        
    '�f�[�^�ŏI�s
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    '�f�[�^����
    Range(Cells(2, "A"), Cells(endRow, "C")).ClearContents
    
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    '�t�H���_�̎擾
    Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
        
    
    For Each f In iFolder.Files: i = i + 1: Next f
    
    If iFolder.Files.Count <> i Then
    
        Set iFolder = Fso.GetFolder(iFolder.ShortPath)
    
    End If
    
    

    
    '�s��
    iRow = 2
    
    '�t�H���_���̃t�@�C��������
    For Each iFile In iFolder.Files
    
    
        Cells(iRow, 1).Value = FolderName & "\" & iFile.Name
        Cells(iRow, 2).Value = FolderName
        Cells(iRow, 3).Value = iFile.Name
        
        iRow = iRow + 1
        
    Next iFile
    
    
    
    
End Sub

'���@�A �R�}���h���C��
Public Sub cmd_�t�@�C�����擾()
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

'���@�B Dir
Public Sub dir_�t�@�C�����擾()

Dim FolderName As String
Dim endRow As Long
Dim FileName As String

Dim i As Long

   '�擾�ꏊ
    FolderName = Range("G1").Value
        
    '�f�[�^�ŏI�s
    endRow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    '�f�[�^����
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

Public Sub �t�@�C�����ύX()  '�����p�X�Ή�
Dim Fso As Object

Dim fileFullPath As String
Dim newfileName As String
Dim FolderName As String
Dim FolderShortName As String
Dim iRow As Long

    iRow = 2
    
    With ThisWorkbook.Sheets("�t�@�C�����ύX")
  
    
        Do
        
             '�ύX�O�̃t���p�X���w��
             fileFullPath = .Cells(iRow, "A").Value
             
             '�ύX��̃t�@�C����
             newfileName = .Cells(iRow, "B").Value
             
             
             Set Fso = CreateObject("Scripting.FileSystemObject")
             
             If Len(fileFullPath) > 259 Then
                
            
                fileFullPath = ChangeShortPath(fileFullPath)
                             
             End If
             
             '�t�@�C������ύX
             Fso.GetFile(fileFullPath).Name = newfileName
             
             '��Еt��
             Set Fso = Nothing
             
         
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     MsgBox "�������܂���"
     
End Sub
