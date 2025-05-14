Attribute VB_Name = "Module1"
Option Explicit
Public Sub DataClassification()

Dim StartRange As Range
Dim TableRange As Range
Dim TargetColumn As Long
Dim mySheet As Worksheet
Dim TableMakeFlag As Boolean
Dim mySlicerCache As SlicerCache
Dim myCriteria As Range
Dim OutRange As Range
Dim OutSheet As Worksheet
Dim i As Long

    Application.ScreenUpdating = False

    With ThisWorkbook.Sheets("�f�[�^")
        
        '�\�̊J�n�Z����ݒ�
        Set StartRange = .Range("A1")
        
        '�\�͈̔͂��擾
        Set TableRange = .Range(StartRange, StartRange.End(xlDown).End(xlToRight))
        
        '���ނ��������ڂ̗�ԍ����w��
        TargetColumn = 3

         '�s�v�ȃV�[�g���폜
        For Each mySheet In ThisWorkbook.Worksheets
            Application.DisplayAlerts = False
            If mySheet.Name <> .Name Then mySheet.Delete
            Application.DisplayAlerts = True
        Next mySheet
        
        '�\���e�[�u���ł͂Ȃ��ꍇ�A�\�͈̔͂��e�[�u���ɕϊ�
        If StartRange.ListObject Is Nothing Then
            .ListObjects.Add(xlSrcRange, TableRange, , xlYes).TableStyle = ""
             TableMakeFlag = True
        End If
        
        '���������쐬�A�������̃Z�����A�N�e�B�u�ɂ���
        .Cells(1, .Cells.SpecialCells(xlLastCell).Column + 2).Value = .Cells(StartRange.Row, TargetColumn).Value
        .Cells(1, .Cells.SpecialCells(xlLastCell).Column + 2).Activate
        
        'SlicerCache���쐬
        Set mySlicerCache = ThisWorkbook.SlicerCaches.Add(StartRange.ListObject, .Cells(StartRange.Row, TargetColumn).Value)
        
        For i = 1 To mySlicerCache.SlicerItems.Count
        
            '��������́i���S��v�Ƃ��邽�߁A�u "'="�v�𓪂ɂ���j
            .Cells(1, Columns.Count).End(xlToLeft).Offset(1, 0).Value = "'=" & mySlicerCache.SlicerItems(i).Value
            
            ' AdvancedFilter���\�b�h�p�̏�����ݒ�
            Set OutSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
            Set myCriteria = .Cells(1, Columns.Count).End(xlToLeft).CurrentRegion
            Set OutRange = OutSheet.Cells(StartRange.Row, StartRange.Column).Resize(1, TableRange.Columns.Count) '�o�͔͈�
            
            '�f�[�^���o�A�ʃV�[�g�֏o��
            TableRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=myCriteria, CopyToRange:=OutRange, Unique:=True
            OutSheet.Cells.EntireColumn.AutoFit
            OutSheet.Name = mySlicerCache.SlicerItems(i).Value
            
        Next i

        mySlicerCache.Delete
        myCriteria.CurrentRegion.Clear

        '�\�͈̔͂��e�[�u���ɕϊ������ꍇ�A�e�[�u����͈͂ɖ߂�
        If TableMakeFlag Then StartRange.ListObject.Unlist
        .Activate
        
    End With
    
    MsgBox "�������܂���"
    
    Application.ScreenUpdating = True
 
End Sub

Public Sub DeleteSheets()
Dim mySheet As Worksheet

  '�s�v�ȃV�[�g���폜
        For Each mySheet In ThisWorkbook.Worksheets
        
            Application.DisplayAlerts = False
            If mySheet.Name <> ThisWorkbook.Sheets("�f�[�^").Name Then mySheet.Delete
            Application.DisplayAlerts = True
        
        Next mySheet
End Sub



