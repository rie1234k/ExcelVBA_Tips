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
'前提条件
    'ファイル内のテーブルに（該当なし）データが入力、連動時に該当するデータがない場合に選択する

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
    
    
    
    'ブック内にあるテーブル(ListObject)とピボットテーブルをコレクション化
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
    
    
    'ブック内のテーブル(ListObject)に「(該当なし)」データを追加
    For i = 1 To TargetListObjects.Count
        
        With TargetListObjects(i)
            
            If WorksheetFunction.CountIf(.ListColumns(1).Range, "(該当なし)") = 0 Then
            
                Set SetRange = .Range.Resize(.Range.Rows.Count + 1, .Range.Columns.Count)
                .Resize SetRange
                .ListRows(.ListRows.Count).Range.Value = "(該当なし)"
            
            End If
            
        End With

    Next i
    
    'テーブルの変更を反映させるため、ピボットテーブル更新
    For i = 1 To TargetPivotTables.Count
        
        TargetPivotTables(i).RefreshTable
    
    Next i
  
    Set TargetListObjects = Nothing
    Set TargetPivotTables = Nothing
    
       
    For Each ChangedSlicer In Target.Slicers
         
        Set ChangedSlicerCache = ChangedSlicer.SlicerCache
        
        '変更したスライサーのオンになっている項目を取得
        Set VisibleChangedItemsDic = CreateObject("Scripting.Dictionary")
                      
        For i = 1 To ChangedSlicerCache.VisibleSlicerItems.Count
            
            SlicerItemName = ChangedSlicerCache.VisibleSlicerItems(i).Name
            VisibleChangedItemsDic.Add SlicerItemName, SlicerItemName
     
        Next i
        
  
        For Each TargetSlicerCache In ThisWorkbook.SlicerCaches
    
            '変更したスライサー以外のソースが同じスライサーに対して処理する
             If TargetSlicerCache.Name <> ChangedSlicerCache.Name _
                And TargetSlicerCache.SourceName = ChangedSlicerCache.SourceName Then
                
                '設定を共通化し、フィルターをクリアする(スライサーの項目の更新も同時に行われる)
                With TargetSlicerCache
                
                    .CrossFilterType = ChangedSlicerCache.CrossFilterType
                    .SortItems = ChangedSlicerCache.SortItems
                    .SortUsingCustomLists = ChangedSlicerCache.SortUsingCustomLists
                    .ShowAllItems = ChangedSlicerCache.ShowAllItems
                    .ClearAllFilters
                
                End With
                
                
                '変更したスライサーにフィルターがかかっている場合、値を連動させる
                If Not ChangedSlicerCache.FilterCleared Then
                        
                    '値を連動させたいスライサーの項目を取得
                    Set TargetItemsDic = CreateObject("Scripting.Dictionary")
                    
                    For i = 1 To TargetSlicerCache.SlicerItems.Count
                        
                        SlicerItemName = TargetSlicerCache.SlicerItems(i).Name
                        TargetItemsDic.Add SlicerItemName, SlicerItemName
                    
                    Next i
                    
                    
                    '値を連動させたいスライサーの項目のうち、変更したスライサーのオンになっている値に該当しない項目はオフにする
                    For i = 0 To TargetItemsDic.Count - 1
                        
                        If Not VisibleChangedItemsDic.Exists(TargetItemsDic.Items()(i)) Then
                            
                            With TargetSlicerCache
                                
                                '最後の1つをオフにする場合には、該当なしをオンにする
                                If .VisibleSlicerItems.Count = 1 Then
                                    
                                     .SlicerItems("(該当なし)").Selected = True
                                     
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


