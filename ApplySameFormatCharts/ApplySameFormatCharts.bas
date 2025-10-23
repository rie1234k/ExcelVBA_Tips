Attribute VB_Name = "Module1"
Option Explicit


Public Sub ApplySameFormatCharts()

Dim TemplateChartObject As ChartObject
Dim TemplateChartName As String
Dim TemplateChartWidth As Double
Dim TemplateChartHeight As Double
Dim TemplateChartType As XlChartType
Dim TemplateAxesFormulas As Collection

Dim Rtn_Size As Long
Dim Rtn_Temp As Long
Dim Rtn_Axes As Long

Dim TargetSheet As Worksheet
Dim TargetChartObject As ChartObject
Dim ChartTitleFormula As String
Dim TargetChartAxes As Axes
Dim ChartAxesFormulas As Collection
Dim i As Long

    If ActiveChart Is Nothing Then
            
        MsgBox "グラフが選択されていません。グラフを選択してから実行してください。"
        Exit Sub
        
    End If
    
    Set TemplateChartObject = ActiveChart.Parent
    
    TemplateChartWidth = TemplateChartObject.Width
    TemplateChartHeight = TemplateChartObject.Height
    TemplateChartType = TemplateChartObject.Chart.ChartType
    TemplateChartName = ThisWorkbook.Path & "\グラフテンプレート" & Format(Now, "yymmdd_hhmmss")
    TemplateChartObject.Chart.SaveChartTemplate TemplateChartName
    
    'テンプレートグラフの軸ラベルを取得
    Set TargetChartAxes = TemplateChartObject.Chart.Axes
    Set TemplateAxesFormulas = New Collection
    i = 1
    Do Until i > TargetChartAxes.Count
        If TargetChartAxes.Item(i).HasTitle Then
            TemplateAxesFormulas.Add TargetChartAxes.Item(i).AxisTitle.Formula, "Item" & i
        Else
            TemplateAxesFormulas.Add "なし", "Item" & i
        End If
        i = i + 1
    Loop

    Rtn_Temp = MsgBox("選択中のグラフの書式を適用しますか？", vbYesNo + vbQuestion)
    Rtn_Size = MsgBox("選択中のグラフの大きさを適用しますか？", vbYesNo + vbQuestion)
    Rtn_Axes = MsgBox("選択中のグラフの軸ラベルを追加しますか？", vbYesNo + vbQuestion)
    
     For Each TargetSheet In ActiveWorkbook.Sheets
      
        For Each TargetChartObject In TargetSheet.ChartObjects
            
            'グラフタイトルを取得
            ChartTitleFormula = ""
            If TargetChartObject.Chart.HasTitle Then
                ChartTitleFormula = TargetChartObject.Chart.ChartTitle.Formula
            End If
            
            '軸ラベルを取得
            Set TargetChartAxes = TargetChartObject.Chart.Axes
            Set ChartAxesFormulas = New Collection
            i = 1
            Do Until i > TargetChartAxes.Count
                If TargetChartAxes.Item(i).HasTitle Then
                    ChartAxesFormulas.Add TargetChartAxes.Item(i).AxisTitle.Formula, "Item" & i
                Else
                    ChartAxesFormulas.Add "なし", "Item" & i
                End If
                i = i + 1
            Loop

            If TargetChartObject.Chart.ChartType = TemplateChartType Then
                
                If Rtn_Temp = vbYes Then
                    
                    TargetChartObject.Chart.ApplyChartTemplate (TemplateChartName)

                    'グラフタイトル数式再設定
                    If ChartTitleFormula <> "" Then
                        TargetChartObject.Chart.ChartTitle.Formula = ChartTitleFormula
                    End If
                    
                    '軸ラベルを再設定
                    i = 1
                    Set TargetChartAxes = TargetChartObject.Chart.Axes
                    
                    Do Until i > TargetChartAxes.Count
                    
                        If TargetChartAxes.Item(i).HasTitle And ChartAxesFormulas("Item" & i) <> "なし" Then
                            TargetChartAxes.Item(i).AxisTitle.Formula = ChartAxesFormulas("Item" & i)
                        Else
                            If TargetChartAxes.Item(i).HasTitle And Rtn_Axes = vbYes Then
                                TargetChartAxes.Item(i).AxisTitle.Formula = TemplateAxesFormulas("Item" & i)
                            Else
                                '軸ラベルを削除
                                Select Case TargetChartAxes.Item(i).Type
                                    Case xlValue
                                        TargetChartObject.Chart.SetElement (msoElementPrimaryValueAxisTitleNone)
                                    Case xlCategory
                                        TargetChartObject.Chart.SetElement (msoElementPrimaryCategoryAxisTitleNone)
                                End Select
                            End If
                        End If
                        
                        i = i + 1
                        
                    Loop
                    
                End If
                
                If Rtn_Size = vbYes Then
                    TargetChartObject.Width = TemplateChartWidth
                    TargetChartObject.Height = TemplateChartHeight
                End If
                
            End If
            
        Next TargetChartObject
        
    Next TargetSheet
    
    Kill TemplateChartName & ".crtx"
  
End Sub

