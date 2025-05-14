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

    With ThisWorkbook.Sheets("データ")
        
        '表の開始セルを設定
        Set StartRange = .Range("A1")
        
        '表の範囲を取得
        Set TableRange = .Range(StartRange, StartRange.End(xlDown).End(xlToRight))
        
        '分類したい項目の列番号を指定
        TargetColumn = 3

         '不要なシートを削除
        For Each mySheet In ThisWorkbook.Worksheets
            Application.DisplayAlerts = False
            If mySheet.Name <> .Name Then mySheet.Delete
            Application.DisplayAlerts = True
        Next mySheet
        
        '表がテーブルではない場合、表の範囲をテーブルに変換
        If StartRange.ListObject Is Nothing Then
            .ListObjects.Add(xlSrcRange, TableRange, , xlYes).TableStyle = ""
             TableMakeFlag = True
        End If
           
        'SlicerCacheを作成
        Set mySlicerCache = ThisWorkbook.SlicerCaches.Add(StartRange.ListObject, .Cells(StartRange.Row, TargetColumn).Value)
        
        For Each TargetItem In mySlicerCache.SlicerItems
            
            'スライサー項目選択（選択したい項目以外の選択を外す）
            mySlicerCache.ClearManualFilter
            For i = 1 To mySlicerCache.SlicerItems.Count
                If mySlicerCache.SlicerItems(i).Value <> TargetItem.Value Then mySlicerCache.SlicerItems(i).Selected = False
            Next i
            
            '新規シートを作成し、コピーして出力
            Set OutSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
            TableRange.SpecialCells(xlCellTypeVisible).Copy OutSheet.Cells(StartRange.Row, StartRange.Column)
            OutSheet.Cells.EntireColumn.AutoFit
            OutSheet.Name = TargetItem.Value
            
        Next TargetItem

        mySlicerCache.Delete
       
        '表の範囲をテーブルに変換した場合、テーブルを範囲に戻す
        If TableMakeFlag Then StartRange.ListObject.Unlist
        .Activate
        
    End With
    
    MsgBox "完了しました"
    
    Application.ScreenUpdating = True
 
End Sub

Public Sub DeleteSheets()
Dim mySheet As Worksheet

  '不要なシートを削除
        For Each mySheet In ThisWorkbook.Worksheets
        
            Application.DisplayAlerts = False
            If mySheet.Name <> ThisWorkbook.Sheets("データ").Name Then mySheet.Delete
            Application.DisplayAlerts = True
        
        Next mySheet
End Sub




