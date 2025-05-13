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
        
        '条件欄を作成、条件欄のセルをアクティブにする
        .Cells(1, .Cells.SpecialCells(xlLastCell).Column + 2).Value = .Cells(StartRange.Row, TargetColumn).Value
        .Cells(1, .Cells.SpecialCells(xlLastCell).Column + 2).Activate
        
        'SlicerCacheを作成
        Set mySlicerCache = ThisWorkbook.SlicerCaches.Add(StartRange.ListObject, .Cells(StartRange.Row, TargetColumn).Value)
        
        For i = 1 To mySlicerCache.SlicerItems.Count
        
            '条件を入力（完全一致とするため、「 "'="」を頭につける）
            .Cells(1, Columns.Count).End(xlToLeft).Offset(1, 0).Value = "'=" & mySlicerCache.SlicerItems(i).Value
            
            ' AdvancedFilterメソッド用の条件を設定
            Set OutSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
            Set myCriteria = .Cells(1, Columns.Count).End(xlToLeft).CurrentRegion
            Set OutRange = OutSheet.Cells(StartRange.Row, StartRange.Column).Resize(1, TableRange.Columns.Count) '出力範囲
            
            'データ抽出、別シートへ出力
            TableRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=myCriteria, CopyToRange:=OutRange, Unique:=True
            OutSheet.Cells.EntireColumn.AutoFit
            OutSheet.Name = mySlicerCache.SlicerItems(i).Value
            
        Next i

        mySlicerCache.Delete
        myCriteria.CurrentRegion.Clear

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



