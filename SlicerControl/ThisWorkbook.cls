VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'�O�����
    '�t�@�C�����̃e�[�u���Ɂi�Y���Ȃ��j�f�[�^�����́A�A�����ɊY������f�[�^���Ȃ��ꍇ�ɑI������

Private Sub Workbook_SheetPivotTableChangeSync(ByVal Sh As Object, ByVal Target As PivotTable)

Dim ChangedSlicer As Slicer
Dim ChangedSlicerCache As SlicerCache
Dim VisibleChangedItemsDic As Object
Dim TargetItemsDic As Object
Dim TargetSlicerCache As SlicerCache
Dim SlicerItemName As String

Dim TargetSheet As Worksheet
Dim TargetListObjects As Collection
Dim TargetPivotTables As Collection
Dim SetRange As Range

Dim i As Long


    If Target.Slicers.Count = 0 Then Exit Sub

    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    
    
    '�u�b�N���ɂ���e�[�u��(ListObject)�ƃs�{�b�g�e�[�u�����R���N�V������
    Set TargetListObjects = New Collection
    Set TargetPivotTables = New Collection
    
    For Each TargetSheet In ThisWorkbook.Worksheets

         For i = 1 To TargetSheet.ListObjects.Count
         
            TargetListObjects.Add TargetSheet.ListObjects(i)
            
         Next i
        
         For i = 1 To TargetSheet.PivotTables.Count
             
             TargetPivotTables.Add TargetSheet.PivotTables(i)
             
         Next i
        
    Next TargetSheet
    
    
    '�u�b�N���̃e�[�u��(ListObject)�Ɂu(�Y���Ȃ�)�v�f�[�^��ǉ�
    For i = 1 To TargetListObjects.Count
        
        With TargetListObjects(i)
            
            If WorksheetFunction.CountIf(.ListColumns(1).Range, "(�Y���Ȃ�)") = 0 Then
            
                Set SetRange = .Range.Resize(.Range.Rows.Count + 1, .Range.Columns.Count)
                .Resize SetRange
                .ListRows(.ListRows.Count).Range.Value = "(�Y���Ȃ�)"
            
            End If
            
        End With

    Next i
    
    '�e�[�u���̕ύX�𔽉f�����邽�߁A�s�{�b�g�e�[�u���X�V
    For i = 1 To TargetPivotTables.Count
        
        TargetPivotTables(i).RefreshTable
    
    Next i
  
    Set TargetListObjects = Nothing
    Set TargetPivotTables = Nothing
    
       
    For Each ChangedSlicer In Target.Slicers
         
        Set ChangedSlicerCache = ChangedSlicer.SlicerCache
        
        '�ύX�����X���C�T�[�̃I���ɂȂ��Ă��鍀�ڂ��擾
        Set VisibleChangedItemsDic = CreateObject("Scripting.Dictionary")
                      
        For i = 1 To ChangedSlicerCache.VisibleSlicerItems.Count
            
            SlicerItemName = ChangedSlicerCache.VisibleSlicerItems(i).Name
            VisibleChangedItemsDic.Add SlicerItemName, SlicerItemName
     
        Next i
        
  
        For Each TargetSlicerCache In ThisWorkbook.SlicerCaches
    
            '�ύX�����X���C�T�[�ȊO�̃\�[�X�������X���C�T�[�ɑ΂��ď�������
             If TargetSlicerCache.Name <> ChangedSlicerCache.Name _
                And TargetSlicerCache.SourceName = ChangedSlicerCache.SourceName Then
                
                '�ݒ�����ʉ����A�t�B���^�[���N���A����(�X���C�T�[�̍��ڂ̍X�V�������ɍs����)
                With TargetSlicerCache
                
                    .CrossFilterType = ChangedSlicerCache.CrossFilterType
                    .SortItems = ChangedSlicerCache.SortItems
                    .SortUsingCustomLists = ChangedSlicerCache.SortUsingCustomLists
                    .ShowAllItems = ChangedSlicerCache.ShowAllItems
                    .ClearAllFilters
                
                End With
                
                
                '�ύX�����X���C�T�[�Ƀt�B���^�[���������Ă���ꍇ�A�l��A��������
                If Not ChangedSlicerCache.FilterCleared Then
                        
                    '�l��A�����������X���C�T�[�̍��ڂ��擾
                    Set TargetItemsDic = CreateObject("Scripting.Dictionary")
                    
                    For i = 1 To TargetSlicerCache.SlicerItems.Count
                        
                        SlicerItemName = TargetSlicerCache.SlicerItems(i).Name
                        TargetItemsDic.Add SlicerItemName, SlicerItemName
                    
                    Next i
                    
                    
                    '�l��A�����������X���C�T�[�̍��ڂ̂����A�ύX�����X���C�T�[�̃I���ɂȂ��Ă���l�ɊY�����Ȃ����ڂ̓I�t�ɂ���
                    For i = 0 To TargetItemsDic.Count - 1
                        
                        If Not VisibleChangedItemsDic.Exists(TargetItemsDic.Items()(i)) Then
                            
                            With TargetSlicerCache
                                
                                '�Ō��1���I�t�ɂ���ꍇ�ɂ́A�Y���Ȃ����I���ɂ���
                                If .VisibleSlicerItems.Count = 1 Then
                                    
                                     .SlicerItems("(�Y���Ȃ�)").Selected = True
                                     
                                End If
                            
                                .SlicerItems(TargetItemsDic.Items()(i)).Selected = False
                            
                            End With
                            
                        End If
                    
                    Next i
                    
                End If
                    
             End If
             
             Set TargetItemsDic = Nothing
             
        Next TargetSlicerCache
         
        Set ChangedSlicerCache = Nothing
        Set VisibleChangedItemsDic = Nothing
        
    Next ChangedSlicer
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

