Attribute VB_Name = "A_Main"
Option Explicit

Public Sub ExtractChartData()

Dim TargetChart As Chart
Dim ExtractChart As Chart
Dim TargetSheet As Worksheet
Dim DataSheet As Worksheet
Dim ExtractSheet As Worksheet
Dim TargetSeries As Series
Dim TargetAxesStr As String
Dim TargetSeriesStr As String
Dim TargetSeriesIndex As String
Dim TargetLabelStr As String
Dim TargetDirection As String
Dim CharStart As Long
Dim CharLength As Long
Dim ChangeFormula As String
Dim LabelRange As Range
Dim SourceRange As Range
Dim ExtractAxesRange As Range
Dim ExtractSeriesRange As Range
Dim ExtractLabelRange As Range
Dim ExtractFormulaRange As Range
Dim i As Long

    If ActiveChart Is Nothing Then
        MsgBox "グラフが選択されていません。グラフを選択してから実行してください。"
        Exit Sub
    End If

    Set TargetChart = ActiveChart
    Set TargetSheet = ActiveSheet

    Set ExtractSheet = ActiveWorkbook.Sheets.Add(After:=Sheets(ActiveWorkbook.Sheets.Count))
    
    With ExtractSheet

        '編集用グラフを出力シートへ移動
        TargetSheet.Activate
        Set ExtractChart = TargetChart.Parent.Duplicate.Chart
        Set ExtractChart = ExtractChart.Location(Where:=xlLocationAsObject, Name:=.Name)
         
        For Each TargetSeries In TargetChart.FullSeriesCollection
             
            'データ範囲が連続している場合のみ処理
            If Len(TargetSeries.FormulaLocal) - Len(Replace(TargetSeries.FormulaLocal, ",", "")) <> 3 Then
                MsgBox "グラフのデータ範囲が連続していないため、抽出できません。"
                Exit Sub
            End If
    
            '系列ラベル範囲取得
            CharStart = InStr(TargetSeries.FormulaLocal, "(") + 1
            CharLength = InStr(CharStart, TargetSeries.FormulaLocal, ",") - CharStart
            TargetLabelStr = Mid(TargetSeries.FormulaLocal, CharStart, CharLength)
            
            '軸ラベル範囲取得
            CharStart = InStr(TargetSeries.FormulaLocal, ",") + 1
            CharLength = InStr(CharStart, TargetSeries.FormulaLocal, ",") - CharStart
            TargetAxesStr = Mid(TargetSeries.FormulaLocal, CharStart, CharLength)
            
            '系列データ範囲取得
            CharStart = InStr(CharStart, TargetSeries.FormulaLocal, ",") + 1
            CharLength = InStr(CharStart, TargetSeries.FormulaLocal, ",") - CharStart
            TargetSeriesStr = Mid(TargetSeries.FormulaLocal, CharStart, CharLength)
            TargetSeriesIndex = Replace(Mid(TargetSeries.FormulaLocal, InStrRev(TargetSeries.FormulaLocal, ",") + 1), ")", "")
            
            Set DataSheet = Range(TargetSeriesStr).Parent
            
            '軸ラベルの方向確認
            If DataSheet.Range(TargetSeriesStr).Columns.Count > DataSheet.Range(TargetSeriesStr).Rows.Count Then
                TargetDirection = "横"
            Else
                TargetDirection = "縦"
            End If
          
            If TargetLabelStr <> "" Then
                Set LabelRange = DataSheet.Range(TargetLabelStr)
            Else
                Set LabelRange = DataSheet.Range("A1") '仮設定
            End If
            
            ChangeFormula = TargetSeries.Formula
            
            '軸データの出力
            If TargetAxesStr <> "" Then
                Set ExtractAxesRange = Switch(TargetDirection = "縦", .Range("A1").Offset(LabelRange.Rows.Count, 0), _
                                             TargetDirection = "横", .Range("A1").Offset(0, LabelRange.Columns.Count))
                Set SourceRange = DataSheet.Range(TargetAxesStr)
                SourceRange.Copy ExtractAxesRange
                Set ExtractAxesRange = ExtractAxesRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
                ExtractAxesRange.Value = SourceRange.Value
                ChangeFormula = Replace(ChangeFormula, TargetAxesStr, ExtractAxesRange.Address(ReferenceStyle:=xlR1C1, External:=True))
            End If
            
            
            '各系列データの出力
            Set SourceRange = DataSheet.Range(TargetSeriesStr)
            Set ExtractSeriesRange = Switch(TargetDirection = "縦", .Cells(LabelRange.Rows.Count + 1, Columns.Count).End(xlToLeft).Offset(0, 1), _
                                             TargetDirection = "横", .Cells(Rows.Count, LabelRange.Columns.Count + 1).End(xlUp).Offset(1, 0))
            SourceRange.Copy ExtractSeriesRange
            Set ExtractSeriesRange = ExtractSeriesRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            ExtractSeriesRange.Value = SourceRange.Value
            ChangeFormula = Replace(ChangeFormula, TargetSeriesStr, ExtractSeriesRange.Address(ReferenceStyle:=xlR1C1, External:=True))
                
            '系列ラベルの出力
            If TargetLabelStr <> "" Then
                Set SourceRange = DataSheet.Range(TargetLabelStr)
                Set ExtractLabelRange = Switch(TargetDirection = "縦", .Cells(LabelRange.Rows.Count + 1, Columns.Count).End(xlToLeft).Offset(-LabelRange.Rows.Count, 0), _
                                             TargetDirection = "横", .Cells(Rows.Count, LabelRange.Columns.Count + 1).End(xlUp).Offset(0, -LabelRange.Columns.Count))
                SourceRange.Copy ExtractLabelRange
                Set ExtractLabelRange = ExtractLabelRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
                ExtractLabelRange.Value = SourceRange.Value
                ChangeFormula = Replace(ChangeFormula, TargetLabelStr, ExtractLabelRange.Address(ReferenceStyle:=xlR1C1, External:=True))
        
            End If
        
            'グラフ範囲の書き換え
            ExtractChart.FullSeriesCollection.Item(TargetSeriesIndex).FormulaR1C1Local = ChangeFormula
        
            '最後にグラフ位置調整
            If TargetSeriesIndex = TargetChart.FullSeriesCollection.Count Then
                Switch(TargetDirection = "縦", .Cells(1, Columns.Count).End(xlToLeft).Offset(5, 2), _
                       TargetDirection = "横", .Cells(Rows.Count, LabelRange.Columns.Count + 1).End(xlUp).Offset(4, 0)).Select
                ExtractChart.Parent.Top = ActiveCell.Top
                ExtractChart.Parent.Left = ActiveCell.Left
            
            End If
            
        Next TargetSeries

        'タイトルが数式の場合、数式で転記
        Set ExtractFormulaRange = .Cells(2, Columns.Count).End(xlToLeft).Offset(-1, 0)
        
        If ExtractChart.HasTitle Then
            If ExtractChart.ChartTitle.Formula <> ExtractChart.ChartTitle.Text Then
               ExtractFormulaRange.Offset(0, 2).Value = "タイトル"
               ExtractFormulaRange.Offset(0, 3).Value = ExtractChart.ChartTitle.Text
               ExtractChart.ChartTitle.Formula = "=" & ExtractFormulaRange.Offset(0, 3).Address(External:=True)
            End If
        End If
        
        '軸ラベルが数式の場合、数式で転記
        i = 1
        Do Until i > ExtractChart.Axes.Count
            If ExtractChart.Axes.Item(i).HasTitle Then
                If ExtractChart.Axes.Item(i).AxisTitle.Formula <> ExtractChart.Axes.Item(i).AxisTitle.Text Then
                    ExtractFormulaRange.Offset(i, 2).Value = "軸ラベル" & i
                    ExtractFormulaRange.Offset(i, 3).Value = ExtractChart.Axes.Item(i).AxisTitle.Text
                    ExtractChart.Axes.Item(i).AxisTitle.Formula = "=" & ExtractFormulaRange.Offset(i, 3).Address(External:=True)
                End If
            End If
            i = i + 1
        Loop
     
     End With

End Sub

