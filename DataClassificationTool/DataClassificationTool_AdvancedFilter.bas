Attribute VB_Name = "Module1"
Option Explicit
Public Sub DataClassification()

Dim StartRange As Range
Dim TableRange As Range
Dim myCriteria As Range
Dim OutRange As Range
Dim TargetSheet As Worksheet
Dim HandlingSheet As Worksheet
Dim OutSheet As Worksheet
Dim mySheet As Worksheet
Dim myListObject As ListObject
Dim mySlicerCache As SlicerCache
Dim CriteriaColumn As Long
Dim i As Long


Const TARGET_COLUMN As Long = 3 '分類したい項目の列番号を指定
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    Set TargetSheet = ThisWorkbook.Sheets("データ")
    
    TargetSheet.Copy Before:=ThisWorkbook.Sheets(1)
    Set HandlingSheet = ThisWorkbook.Sheets(1)
    HandlingSheet.Name = "処理用"
    
    With HandlingSheet
        
        '表の開始セルを設定
        Set StartRange = .Range("A1")
        
        '表の範囲を取得
        Set TableRange = .Range(StartRange, StartRange.End(xlDown).End(xlToRight))
        
        '不要なシートを削除
        Application.DisplayAlerts = False
        For Each mySheet In ThisWorkbook.Worksheets
            If mySheet.Name Like "分類*" Then mySheet.Delete
        Next mySheet
        Application.DisplayAlerts = True
        
        'テーブル化されている場合、いったん範囲に戻して、再度テーブル化して確実にテーブルを取得する
        If Not StartRange.ListObject Is Nothing Then
            StartRange.ListObject.Unlist
        End If
        
        Set myListObject = .ListObjects.Add(xlSrcRange, TableRange, , xlYes)
        
        '条件欄を作成
        CriteriaColumn = StartRange.Column + TableRange.Columns.Count + 2
        .Cells(1, CriteriaColumn).Value = .Cells(StartRange.Row, TARGET_COLUMN).Value
        '条件欄のセルをアクティブにする(AdvancedFilter実行時にテーブル内のセルがアクティブの場合、エラーになるため）
        .Cells(1, CriteriaColumn).Activate
        
        'SlicerCacheを作成
        Set mySlicerCache = ThisWorkbook.SlicerCaches.Add(myListObject, .Cells(StartRange.Row, TARGET_COLUMN).Value)
        
        For i = 1 To mySlicerCache.SlicerItems.Count
        
            '条件を入力（完全一致とするため、「 "'="」を頭につける）
            .Cells(2, CriteriaColumn).Value = "'=" & mySlicerCache.SlicerItems(i).Value
            
            ' AdvancedFilterメソッド用の条件を設定
            Set OutSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
            Set myCriteria = .Cells(1, CriteriaColumn).CurrentRegion
            Set OutRange = OutSheet.Cells(StartRange.Row, StartRange.Column).Resize(1, TableRange.Columns.Count) '出力範囲
            
            'データ抽出、別シートへ出力、「Unique:=True」で重複データを削除
            TableRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=myCriteria, CopyToRange:=OutRange, Unique:=True
            OutSheet.Cells.EntireColumn.AutoFit
            OutSheet.Name = "分類" & i
            
        Next i

        mySlicerCache.Delete
    
        Application.DisplayAlerts = False
        .Delete
        Application.DisplayAlerts = True
        
    End With
    
    TargetSheet.Activate
    MsgBox "完了しました"
    
CleanUp:
    
    Set myListObject = Nothing
    Set mySlicerCache = Nothing
    Set StartRange = Nothing
    Set TableRange = Nothing
    Set myCriteria = Nothing
    Set OutRange = Nothing
    Set mySheet = Nothing
    Set OutSheet = Nothing
    Set HandlingSheet = Nothing
    Set TargetSheet = Nothing
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    
    MsgBox "エラーが発生したため、処理を中止します。"
    
    Application.DisplayAlerts = False
    For Each mySheet In ThisWorkbook.Worksheets
        If mySheet.Name = "処理用" Or mySheet.Name Like "分類*" Then mySheet.Delete
    Next mySheet
    Application.DisplayAlerts = True
    
    GoTo CleanUp
 
End Sub

Public Sub DeleteSheets()
Dim mySheet As Worksheet

  '不要なシートを削除
    Application.DisplayAlerts = False
    For Each mySheet In ThisWorkbook.Worksheets
        If mySheet.Name Like "分類*" Then mySheet.Delete
    Next mySheet
    Application.DisplayAlerts = True
        
End Sub



