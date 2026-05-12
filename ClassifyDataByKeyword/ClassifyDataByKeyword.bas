Attribute VB_Name = "Module1"
Option Explicit
Public Sub ClassifyDataByKeyword()

Dim iCount As Long
Dim endRow As Long
Dim SearchWordSheet As Worksheet
Dim TargetSheet As Worksheet
Dim SearchList As Variant
Dim TargetRange As Range
    
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Set SearchWordSheet = ThisWorkbook.Sheets("検索ワード対応表")
    Set TargetSheet = ThisWorkbook.Sheets("データ処理")
    
    With SearchWordSheet
        endRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        'A列：検索ワード　B列：入力値
        SearchList = SearchWordSheet.Range(.Range("A2"), .Cells(endRow, "B")).Value
    End With
    
    With TargetSheet
        
        'オートフィルターが設定されている場合には解除
        If .AutoFilterMode Then .AutoFilterMode = False
    
        '非表示セルを表示
        .Cells.EntireRow.Hidden = False
        
        '入力データをクリア
        .Range("A1").CurrentRegion.Offset(1, 1).ClearContents

        endRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        For iCount = 1 To UBound(SearchList, 1)
            
            Application.StatusBar = "検索ワード " & iCount & "/" & UBound(SearchList, 1) & " 件処理中..."
            
            .Range("A1").AutoFilter Field:=1, Criteria1:="*" & SearchList(iCount, 1) & "*"
            
            Set TargetRange = Nothing
            On Error Resume Next
            Set TargetRange = .Range(.Range("B2"), .Cells(endRow, "B")).SpecialCells(xlCellTypeVisible)
            On Error GoTo ErrHandler
            
            If Not TargetRange Is Nothing Then
                TargetRange.Value = SearchList(iCount, 2)
            End If
            
            'オートフィルターの条件を解除
            If .AutoFilterMode Then
                If .FilterMode Then .AutoFilter.ShowAllData
            End If
            
        Next iCount
        
    End With
    
Cleanup:

    If TargetSheet.AutoFilterMode Then TargetSheet.AutoFilterMode = False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & "( " & Err.Description & ")"
    GoTo Cleanup
    
End Sub

