Attribute VB_Name = "Module1"
Option Explicit

Public Sub AutofilterExcludeMultiple()

Dim TargetColumnNo As Long
Dim ItemCount As Long
Dim TargetItemArray() As String
Dim i As Long

    '項目列番号・除外対象項目を取得
    With ThisWorkbook.Sheets("除外条件")
    
        TargetColumnNo = WorksheetFunction.Match(.Range("A2").Value, ActiveSheet.Rows(1), 0)
        
        ItemCount = .Range(.Range("C2"), .Range("C2").End(xlDown)).Count
        
        ReDim TargetItemArray(ItemCount - 1)
        
        For i = 0 To ItemCount - 1
        
            TargetItemArray(i) = .Cells(i + 2, "C").Value
            
        Next i
        
        
    End With


     With ActiveSheet
        
        'オートフィルターが設定されている場合には解除
        If Not .AutoFilter Is Nothing Then .Range("A1").AutoFilter
        
        '塗りつぶし解除
        .Columns(1).Interior.Color = xlNone
        
        
        '除外対象で絞り込む
        .Range("A1").AutoFilter Field:=TargetColumnNo, Criteria1:=TargetItemArray, Operator:=xlFilterValues

        '除外対象のA列のセルを塗りつぶし
        .Range(.Range("A2"), .Range("A2").End(xlDown)).Interior.Color = vbYellow
        
        .ShowAllData
    
        '塗りつぶされていないセルを抽出 ＝ 除外対象以外のデータ
         .Range("A1").AutoFilter Field:=1, Operator:=xlFilterNoFill
            
    End With
    
    

End Sub


Public Sub DataExtract()

Dim TargetColumnNo As Long
Dim ItemCount As Long
Dim TargetItemArray() As String
Dim i As Long

       '項目列番号・除外対象項目を取得
    With ThisWorkbook.Sheets("除外条件")
    
        TargetColumnNo = WorksheetFunction.Match(.Range("A2").Value, ActiveSheet.Rows(1), 0)
        
        ItemCount = .Range(.Range("C2"), .Range("C2").End(xlDown)).Count
        
        ReDim TargetItemArray(ItemCount - 1)
        
        For i = 0 To ItemCount - 1
        
            TargetItemArray(i) = .Cells(i + 2, "C").Value
            
        Next i
        
        
    End With
    
    
    With ActiveSheet
        
        '------- 初期化 -------
        If Not .AutoFilter Is Nothing Then .Range("A1").AutoFilter
        .Columns(1).Interior.Color = xlNone
        
        
        '------- 抽出したいデータに印をつける -------
        .Range("A1").AutoFilter Field:=TargetColumnNo, Criteria1:=TargetItemArray, Operator:=xlFilterValues
        .Range(.Range("A2"), .Range("A2").End(xlDown)).Interior.Color = vbYellow
        .ShowAllData
        
        
        '------- その他のデータを削除 -------
        .Range("A1").AutoFilter Field:=1, Operator:=xlFilterNoFill
        .Range("A1").CurrentRegion.Offset(1).EntireRow.Delete
        .Range("A1").AutoFilter
        .Columns(1).Interior.Color = xlNone
         
    End With

End Sub
