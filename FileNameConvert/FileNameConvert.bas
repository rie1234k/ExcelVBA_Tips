Attribute VB_Name = "FileNameConvert"
Option Explicit

Public Sub GetFileName()

Dim FolderName As String
Dim Fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long
 
    With ThisWorkbook.Sheets("�t�@�C�����擾")
 
        '�擾�ꏊ
        FolderName = .Range("G1").Value
              
        '�f�[�^����
        .Range("A1").CurrentRegion.Offset(1).ClearContents
  
        Set Fso = CreateObject("Scripting.FileSystemObject")
        
        '�t�H���_�̎擾
        Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
       
        '�s��
        iRow = 2
        
        '�t�H���_���̃t�@�C��������
        For Each iFile In iFolder.Files
        
            .Cells(iRow, 1).Value = FolderName & "\" & iFile.Name
            .Cells(iRow, 2).Value = FolderName
            .Cells(iRow, 3).Value = iFile.Name
            
            iRow = iRow + 1
            
        Next iFile
    
    End With
    
    Set Fso = Nothing
        
End Sub

Public Sub ChangeFileName()

Dim Fso As Object
Dim FileFullPath As String
Dim newFileName As String
Dim iRow As Long

    iRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
     
    With ThisWorkbook.Sheets("�t�@�C�����ύX")
      
        Do
        
             '�ύX�O�̃t���p�X���w��
             FileFullPath = .Cells(iRow, "A").Value
             
             '�ύX��̃t�@�C����
             newFileName = .Cells(iRow, "B").Value
             
             '�V���[�g�p�X�ɕϊ�
             FileFullPath = ChangeShortPath(FileFullPath)

             '�t�@�C������ύX
             Fso.GetFile(FileFullPath).Name = newFileName
            
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     Set Fso = Nothing
     
     MsgBox "�������܂���"
     
End Sub
