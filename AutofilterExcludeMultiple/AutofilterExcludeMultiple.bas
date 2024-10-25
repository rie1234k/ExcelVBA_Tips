Attribute VB_Name = "Module1"
Option Explicit

Public Sub AutofilterExcludeMultiple()

Dim TargetColumnNo As Long
Dim TargetItemString As String
Dim TargetItemArray As Variant

    '���ڗ�ԍ��E���O�Ώۍ��ڂ��擾
    With ThisWorkbook.Sheets("���O����")
    
        TargetColumnNo = WorksheetFunction.Match(.Range("A2").Value, ActiveSheet.Rows(1), 0)
        TargetItemString = WorksheetFunction.TextJoin(",", True, .Range(.Range("C2"), .Range("C2").End(xlDown)))
        TargetItemArray = Split(TargetItemString, ",")
    
    End With


     With ActiveSheet
        
        '�I�[�g�t�B���^�[����
        If Not .AutoFilter Is Nothing Then .Range("A1").AutoFilter
        
        '�h��Ԃ�����
        .Columns(1).Interior.Color = xlNone
        
        
        '���O�Ώۂōi�荞��
        .Range("A1").AutoFilter Field:=TargetColumnNo, Criteria1:=TargetItemArray, Operator:=xlFilterValues

        '���O�Ώۂ�A��̃Z����h��Ԃ�
        .Range(.Range("A2"), .Range("A2").End(xlDown)).Interior.Color = vbYellow
        
        .Range("A1").AutoFilter
        '�h��Ԃ���Ă��Ȃ��Z���𒊏o �� ���O�ΏۈȊO�̃f�[�^
         .Range("A1").AutoFilter Field:=1, Operator:=xlFilterNoFill
            
    End With
    
    

End Sub


Public Sub DataExtract()

Dim TargetColumnNo As Long
Dim TargetItemString As String
Dim TargetItemArray As Variant

     '���ڗ�ԍ��E�Ώۍ��ڂ��擾
    With ThisWorkbook.Sheets("���O����")
    
        TargetColumnNo = WorksheetFunction.Match(.Range("A2").Value, ActiveSheet.Rows(1), 0)
        TargetItemString = WorksheetFunction.TextJoin(",", True, .Range(.Range("C2"), .Range("C2").End(xlDown)))
        TargetItemArray = Split(TargetItemString, ",")
    
    End With
    
    
    With ActiveSheet
        
        '------- ������ -------
        If Not .AutoFilter Is Nothing Then .Range("A1").AutoFilter
        .Columns(1).Interior.Color = xlNone
        
        
        '------- ���o�������f�[�^�Ɉ������ -------
        .Range("A1").AutoFilter Field:=TargetColumnNo, Criteria1:=TargetItemArray, Operator:=xlFilterValues
        .Range(.Range("A2"), .Range("A2").End(xlDown)).Interior.Color = vbYellow
        .ShowAllData
        
        
        '------- ���̑��̃f�[�^���폜 -------
        .Range("A1").AutoFilter Field:=1, Operator:=xlFilterNoFill
        .Range("A1").CurrentRegion.Offset(1).EntireRow.Delete
        .Range("A1").AutoFilter
        .Columns(1).Interior.Color = xlNone
         
    End With

End Sub
