Attribute VB_Name = "Module1"
Option Explicit
Public Sub ClassifyDataByKeyword()

Dim MapSheet As Worksheet
Dim DataSheet As Worksheet
Dim TargetData As Variant
Dim ResultData As Variant
Dim SearchList As Variant
Dim DataEndRow As Long
Dim MapEndRow As Long
Dim i As Long
Dim j As Long
Dim temp1 As String, temp2 As String, temp3 As Long
Dim Data As String
Dim Key As String

    Set MapSheet = ThisWorkbook.Sheets("検索ワード対応表")
    Set DataSheet = ThisWorkbook.Sheets("データ処理")
    
    With DataSheet
         
         'オートフィルターが設定されている場合には解除
        If .AutoFilterMode Then .AutoFilterMode = False
        '非表示セルを表示
        .Cells.EntireRow.Hidden = False
        '入力データをクリア
        .Range("A1").CurrentRegion.Offset(1, 1).ClearContents
        
        DataEndRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        TargetData = .Range(.Range("A2"), .Cells(DataEndRow, "A")).Value
        ReDim ResultData(1 To UBound(TargetData, 1), 1 To 2)
        
    End With
    
     With MapSheet
        MapEndRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        'A列：検索ワード　B列：入力値　C列：優先順位
        SearchList = .Range(.Range("A2"), .Cells(MapEndRow, "C")).Value
    End With
    
    For i = 1 To UBound(SearchList, 1)
        For j = i + 1 To UBound(SearchList, 1)
            If SearchList(i, 3) > SearchList(j, 3) _
               Or (SearchList(i, 3) = SearchList(j, 3) And Len(SearchList(i, 1)) < Len(SearchList(j, 1))) Then
                temp1 = SearchList(i, 1)
                temp2 = SearchList(i, 2)
                temp3 = SearchList(i, 3)
                SearchList(i, 1) = SearchList(j, 1)
                SearchList(i, 2) = SearchList(j, 2)
                SearchList(i, 3) = SearchList(j, 3)
                SearchList(j, 1) = temp1
                SearchList(j, 2) = temp2
                SearchList(j, 3) = temp3
            End If
        Next j
    Next i

    For i = 1 To UBound(SearchList, 1)
        Key = "*" & LCase(Trim(SearchList(i, 1))) & "*"
        For j = 1 To UBound(TargetData, 1)
            If TargetData(j, 1) <> "" Then
                Data = LCase(Trim(TargetData(j, 1)))
                If ResultData(j, 1) = "" Then
                    If Data Like Key Then
                        ResultData(j, 1) = SearchList(i, 2)
                        If Data = Key Then
                            ResultData(j, 2) = "完全一致：" & SearchList(i, 1) & "（優先度：" & SearchList(i, 3) & "）"
                        Else
                            ResultData(j, 2) = "部分一致：" & SearchList(i, 1) & "（優先度：" & SearchList(i, 3) & "）"
                        End If
                    End If
                End If
            End If
        Next j

    Next i
    
    With DataSheet
        .Range(.Range("B2"), .Cells(DataEndRow, "C")).Value = ResultData
    End With
    
End Sub

