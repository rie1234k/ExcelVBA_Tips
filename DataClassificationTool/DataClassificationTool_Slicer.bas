Attribute VB_Name = "Module1"
Option Explicit

Public Sub DataClassification_Slicer()

Dim StartRange As Range
Dim TableRange As Range
Dim TargetColumn As Long
Dim mySheet As Worksheet
Dim TableMakeFlag As Boolean
Dim mySlicerCache As SlicerCache
Dim TargetItem As SlicerItem
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
           
        'SlicerCache���쐬
        Set mySlicerCache = ThisWorkbook.SlicerCaches.Add(StartRange.ListObject, .Cells(StartRange.Row, TargetColumn).Value)
        
        For Each TargetItem In mySlicerCache.SlicerItems
            
            '�X���C�T�[���ڑI���i�I�����������ڈȊO�̑I�����O���j
            mySlicerCache.ClearManualFilter
            For i = 1 To mySlicerCache.SlicerItems.Count
                If mySlicerCache.SlicerItems(i).Value <> TargetItem.Value Then mySlicerCache.SlicerItems(i).Selected = False
            Next i
            
            '�V�K�V�[�g���쐬���A�R�s�[���ďo��
            Set OutSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
            TableRange.SpecialCells(xlCellTypeVisible).Copy OutSheet.Cells(StartRange.Row, StartRange.Column)
            OutSheet.Cells.EntireColumn.AutoFit
            OutSheet.Name = TargetItem.Value
            
        Next TargetItem

        mySlicerCache.Delete
       
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




