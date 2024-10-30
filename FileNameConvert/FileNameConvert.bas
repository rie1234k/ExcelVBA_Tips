Attribute VB_Name = "FileNameConvert"
Option Explicit

'���@�@ FileSystemObject �����p�X�ɑΉ�
Public Sub fso_GetFileName()

Dim FolderName As String
Dim Fso As Object
Dim iFolder As Object
Dim iFile As Object
Dim iRow As Long

Dim i As Long
Dim f As Object

     
    With ThisWorkbook.Sheets("�t�@�C�����擾")
 
        '�擾�ꏊ
        FolderName = .Range("G1").Value
              
        '�f�[�^����
        .Range("A1").CurrentRegion.Offset(1).ClearContents

        
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
        
        
            .Cells(iRow, 1).Value = FolderName & "\" & iFile.Name
            .Cells(iRow, 2).Value = FolderName
            .Cells(iRow, 3).Value = iFile.Name
            
            iRow = iRow + 1
            
        Next iFile
        
    
    End With
    
    Set Fso = Nothing
        
End Sub


Public Sub ChangeFileName()  '�����p�X�Ή�
Dim Fso As Object

Dim FileFullPath As String
Dim newfileName As String
Dim FolderName As String
Dim FolderShortName As String
Dim iRow As Long

    iRow = 2
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
     
    With ThisWorkbook.Sheets("�t�@�C�����ύX")
      
        Do
        
             '�ύX�O�̃t���p�X���w��
             FileFullPath = .Cells(iRow, "A").Value
             
             '�ύX��̃t�@�C����
             newfileName = .Cells(iRow, "B").Value

             FileFullPath = ChangeShortPath(FileFullPath)

             '�t�@�C������ύX
             Fso.GetFile(FileFullPath).Name = newfileName
            
             iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
     
     Set Fso = Nothing
     
     MsgBox "�������܂���"
     
End Sub
