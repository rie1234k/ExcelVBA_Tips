Attribute VB_Name = "Module1"
Option Explicit

Public Sub DataClassification_Slicer()

Dim StartRange As Range
Dim TableRange As Range
Dim CopyRange As Range
Dim TargetSheet As Worksheet
Dim HandlingSheet As Worksheet
Dim mySheet As Worksheet
Dim myListObject As ListObject
Dim mySlicerCache As SlicerCache
Dim TargetItem As SlicerItem
Dim OutSheet As Worksheet
Dim i As Long
Dim iCount As Long

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
        myListObject.TableStyle = ""
        
        'SlicerCacheを作成
        Set mySlicerCache = ThisWorkbook.SlicerCaches.Add(myListObject, .Cells(StartRange.Row, TARGET_COLUMN).Value)
        
        iCount = 1
        
        For Each TargetItem In mySlicerCache.SlicerItems
            
            'スライサー項目選択（選択したい項目以外の選択を外す）
            mySlicerCache.ClearManualFilter
            For i = 1 To mySlicerCache.SlicerItems.Count
                If mySlicerCache.SlicerItems(i).Value <> TargetItem.Value Then mySlicerCache.SlicerItems(i).Selected = False
            Next i
            
            '新規シートを作成し、コピーして出力
            Set CopyRange = Nothing
            On Error Resume Next
            Set CopyRange = TableRange.SpecialCells(xlCellTypeVisible)
            On Error GoTo ErrHandler
            
            If Not CopyRange Is Nothing Then
                Set OutSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
                CopyRange.Copy OutSheet.Cells(StartRange.Row, StartRange.Column)
                OutSheet.Cells.EntireColumn.AutoFit
                OutSheet.Name = "分類" & iCount
                iCount = iCount + 1
            End If
            
        Next TargetItem
        
        mySlicerCache.Delete

        Application.DisplayAlerts = False
        .Delete
        Application.DisplayAlerts = True
        
    End With

    TargetSheet.Activate
    
    MsgBox "完了しました"

CleanUp:

    Set TargetItem = Nothing
    Set mySlicerCache = Nothing
    Set myListObject = Nothing
    Set StartRange = Nothing
    Set TableRange = Nothing
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
        For Each mySheet In ThisWorkbook.Worksheets
            Application.DisplayAlerts = False
            If mySheet.Name Like "分類*" Then mySheet.Delete
            Application.DisplayAlerts = True
        Next mySheet
        
End Sub




