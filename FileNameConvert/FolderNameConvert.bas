Attribute VB_Name = "FolderNameConvert"
Option Explicit

Public Sub GetFolderName()

Dim FolderName As String
Dim Fso As Object
Dim iFolder As Object
Dim subFolder As Object
Dim iRow As Long

    With ThisWorkbook.Sheets("�t�H���_���擾")
  
        '�擾�ꏊ
        FolderName = .Range("G1").Value
            
        '�f�[�^����
        .Range("A1").CurrentRegion.Offset(1).ClearContents
        
        Set Fso = CreateObject("Scripting.FileSystemObject")
        
        '�t�H���_�̎擾
        Set iFolder = Fso.GetFolder(ChangeShortPath(FolderName))
                
        '�s��
        iRow = 2
        
        '�t�H���_���̃T�u�t�H���_������
        For Each subFolder In iFolder.SubFolders
        
            .Cells(iRow, 1).Value = FolderName & "\" & subFolder.Name
            .Cells(iRow, 2).Value = subFolder.Name
            
            iRow = iRow + 1
            
        Next subFolder
        
    End With
    
    Set Fso = Nothing
    
    
End Sub

Public Sub ChangeFolderName()

Dim Fso As Object
Dim FolderFullPath As String
Dim newFolderName As String
Dim iRow As Long

    iRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
             
    With ThisWorkbook.Sheets("�t�H���_���ύX")
  
        Do
        
            '�ύX�O�̃t�H���_�̃p�X���w��
            FolderFullPath = .Cells(iRow, "A").Value
            
            '�ύX��̃t�H���_��
            newFolderName = .Cells(iRow, "B").Value
    
            '�V���[�g�p�X�ɕϊ�
            FolderFullPath = ChangeShortPath(FolderFullPath)
            
            '�t�H���_����ύX
            Fso.GetFolder(FolderFullPath).Name = newFolderName
        
            iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
    
    Set Fso = Nothing

    MsgBox "�������܂���"
    
End Sub

