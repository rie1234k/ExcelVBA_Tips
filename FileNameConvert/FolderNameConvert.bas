Attribute VB_Name = "FolderNameConvert"
Option Explicit

'���@�@ FileSystemObject�@�����p�X�ɑΉ�
Public Sub fso_GetFolderName()
Dim FolderName As String
Dim endRow As Long
Dim Fso As Object
Dim iFolder As Object
Dim subFolder As Object
Dim iRow As Long
Dim i As Long
Dim f As Object


    With ThisWorkbook.Sheets("�t�H���_���擾")
  
        '�擾�ꏊ
        FolderName = .Range("G1").Value
            
        '�ŏI�s
        endRow = .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
        
        '�f�[�^����
        .Range(.Cells(2, "A"), .Cells(endRow, "B")).ClearContents
        
        
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
        
        
            .Cells(iRow, 1).Value = FolderName & "\" & subFolder.Name
            .Cells(iRow, 2).Value = subFolder.Name
            
            iRow = iRow + 1
            
        Next subFolder
        
        
    End With
    
    Set Fso = Nothing
    
    
End Sub


Public Sub ChangeFolderName()  '�����p�X�Ή�

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
    
           
           FolderFullPath = ChangeShortPath(FolderFullPath)
            
            '�t�H���_����ύX
            Fso.GetFolder(FolderFullPath).Name = newFolderName
        
            iRow = iRow + 1
             
         Loop Until .Cells(iRow, "A").Value = ""
    
    End With
    
    
    MsgBox "�������܂���"
    
    Set Fso = Nothing
    
End Sub

