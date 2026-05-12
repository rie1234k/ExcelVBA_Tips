Attribute VB_Name = "Module1"
Option Explicit

Public Sub AutofilterExcludeMultiple()

Dim TargetCol As Long
Dim PaintCol As Long
Dim ItemCount As Long
Dim TargetItemArray() As String
Dim i As Long
Dim LastCol As Long
Dim LastRow As Long

    PaintCol = 1 '塗りつぶし作業列をA列とする
    
    '項目列番号・除外対象項目を取得
    With ThisWorkbook.Sheets("除外条件")
        
        TargetCol = WorksheetFunction.Match(.Range("A2").Value, ActiveSheet.Rows(1), 0)
        ItemCount = .Range(.Range("C2"), .Range("C2").End(xlDown)).Count
        
        ReDim TargetItemArray(ItemCount - 1)
        For i = 0 To ItemCount - 1
            TargetItemArray(i) = .Cells(i + 2, "C").Value
        Next i
        
    End With

     With ActiveSheet
        
        'オートフィルターが設定されている場合には解除
        If .AutoFilterMode Then .AutoFilterMode = False
        
        '塗りつぶし解除
        .Columns(PaintCol).Interior.Color = xlNone

        '除外対象で絞り込む
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        LastRow = .Cells(.Rows.Count, TargetCol).End(xlUp).Row
        .Range(.Range("A1"), .Cells(LastRow, LastCol)).AutoFilter Field:=TargetCol, Criteria1:=TargetItemArray, Operator:=xlFilterValues

        '除外対象のA列のセルを塗りつぶし
        .Range(.Cells(2, PaintCol), .Cells(.Rows.Count, PaintCol).End(xlUp)).SpecialCells(xlCellTypeVisible).Interior.Color = vbYellow
        
        .AutoFilter.ShowAllData
    
        '塗りつぶされていないセルを抽出 ＝ 除外対象以外のデータ
         .Range(.Range("A1"), .Cells(LastRow, LastCol)).AutoFilter Field:=PaintCol, Operator:=xlFilterNoFill
            
    End With

End Sub


Public Sub DataExtract()

Dim TargetCol As Long
Dim PaintCol As Long
Dim ItemCount As Long
Dim TargetItemArray() As String
Dim i As Long
Dim LastCol As Long
Dim LastRow As Long

    PaintCol = 1 '塗りつぶし作業列をA列とする
    
    '項目列番号・除外対象項目を取得
    With ThisWorkbook.Sheets("除外条件")
        TargetCol = WorksheetFunction.Match(.Range("A2").Value, ActiveSheet.Rows(1), 0)
        ItemCount = .Range(.Range("C2"), .Range("C2").End(xlDown)).Count
        
        ReDim TargetItemArray(ItemCount - 1)
        For i = 0 To ItemCount - 1
            TargetItemArray(i) = .Cells(i + 2, "C").Value
        Next i
    End With

     With ActiveSheet
        '------- 初期化 -------
        If .AutoFilterMode Then
            If .FilterMode Then .AutoFilter.ShowAllData
        End If
        .Columns(PaintCol).Interior.Color = xlNone

        '------- 抽出したいデータに印をつける -------
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        LastRow = .Cells(.Rows.Count, TargetCol).End(xlUp).Row
        .Range(.Range("A1"), .Cells(LastRow, LastCol)).AutoFilter Field:=TargetCol, Criteria1:=TargetItemArray, Operator:=xlFilterValues
        .Range(.Cells(2, PaintCol), .Cells(.Rows.Count, PaintCol).End(xlUp)).SpecialCells(xlCellTypeVisible).Interior.Color = vbYellow
        .AutoFilter.ShowAllData
    
        '------- その他のデータを削除 -------
        .Range(.Range("A1"), .Cells(LastRow, LastCol)).AutoFilter Field:=PaintCol, Operator:=xlFilterNoFill
        .Range(.Range("A1"), .Cells(LastRow, LastCol)).Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If
        .Columns(PaintCol).Interior.Color = xlNone
     End With
     
End Sub
