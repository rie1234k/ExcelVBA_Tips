Attribute VB_Name = "ShortPath"
Option Explicit

Public Function ChangeShortPath(FullPath As String) As String

Dim Fso As Object
Dim TargetPath As String
Dim LastPath As String

Dim i As Long

   Set Fso = CreateObject("Scripting.FileSystemObject")
        
    TargetPath = FullPath
    
    '���݂���t�H���_�܂ők���āA���݂���t�H���_���V���[�g�p�X�ɕϊ�����
    Do Until Fso.FolderExists(TargetPath)
    
        'GetFileName�́A�t�@�C���Ɍ��炸�A�ŏI�v�f�Ǝ��o��
        LastPath = "\" & Fso.GetFileName(TargetPath) & LastPath
        
        TargetPath = Fso.GetParentFolderName(TargetPath)
        
        If TargetPath = "" Then
            
            MsgBox "�w�肳�ꂽ�t�H���_�͑��݂��܂���B"
            End
            
        End If
        
    Loop
    
    TargetPath = Fso.GetFolder(TargetPath).ShortPath & LastPath
    
    ChangeShortPath = TargetPath
        
    If Fso.FileExists(ChangeShortPath) = False And Fso.FolderExists(ChangeShortPath) = False Then
        
        MsgBox "�w�肳�ꂽ�p�X�͑��݂��܂���B"
        End
        
    End If
    
    Set Fso = Nothing
    
End Function
