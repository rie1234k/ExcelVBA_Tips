Attribute VB_Name = "Module1"
Option Explicit
Public Sub ClassifyDataByKeyword()

Dim iCount As Long
Dim endRow As Long
Dim TargetSheet As Worksheet
Dim SearchWordSheet As Worksheet
Dim SearchList As Variant
Dim TargetRange As Range
    
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Set TargetSheet = ThisWorkbook.Sheets("データ処理")
    Set SearchWordSheet = ThisWorkbook.Sheets("検索ワード対応表")
    
    'オートフィルターが設定されている場合には解除
    If TargetSheet.AutoFilterMode Then TargetSheet.AutoFilterMode = False
    
    '非表示セルを表示
    TargetSheet.Cells.EntireRow.Hidden = False
    
    '入力データをクリア
    TargetSheet.Range("A1").CurrentRegion.Offset(1, 1).ClearContents

    With SearchWordSheet
        endRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        SearchList = SearchWordSheet.Range(.Range("A2"), .Cells(endRow, "B")).Value
    End With
    
    With TargetSheet
    
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
            
            'オートフィルターを解除
            If .AutoFilterMode Then .AutoFilterMode = False

        Next iCount
        
    End With
    
Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & "( " & Err.Description & ")"
    GoTo Cleanup
    
End Sub

