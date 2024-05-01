Attribute VB_Name = "ShortPath"
Option Explicit

Public Function ChangeShortPath(FullPath As String) As String

Dim Fso As Object
Dim TargetFolder As String
Dim endPath As String
Dim FolderPath As String

    endPath = ""
    TargetFolder = ""
    FolderPath = ""

    Set Fso = CreateObject("Scripting.FileSystemObject")
        
        
    TargetFolder = FullPath
    
    '���݂���t�H���_�܂ők���āA���݂���t�H���_���V���[�g�p�X�ɕϊ�����
    Do Until Fso.FolderExists(TargetFolder)
        
        If endPath <> "" Then
        
            endPath = Fso.GetFileName(TargetFolder) & "\" & endPath
        
        Else
        
            endPath = Fso.GetFileName(TargetFolder)
        
        End If
        
        TargetFolder = Fso.GetParentFolderName(TargetFolder)
        
    
    Loop
    
    
    If endPath <> "" Then
        
        TargetFolder = Fso.GetFolder(TargetFolder).ShortPath & "\" & endPath
         
    End If
    
    
    '�ϊ������V���[�g�p�X�Ɩ����̃p�X���q�����p�X�ɂ��āA�����ꍇ�͍ēx�V���[�g�p�X�ɕϊ�����
    endPath = ""
    
    Do Until Fso.FolderExists(TargetFolder)
        
        If endPath <> "" Then
        
            endPath = Fso.GetFileName(TargetFolder) & "\" & endPath
        
        Else
        
            endPath = Fso.GetFileName(TargetFolder)
        
        End If
        
        TargetFolder = Fso.GetParentFolderName(TargetFolder)
        
    
    Loop
    
    
    If endPath <> "" Then
        
        ChangeShortPath = Fso.GetFolder(TargetFolder).ShortPath & "\" & endPath
    
    Else
    
         ChangeShortPath = TargetFolder
         
    End If
    

End Function


