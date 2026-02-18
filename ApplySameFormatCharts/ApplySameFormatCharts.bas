Attribute VB_Name = "Module1"
Option Explicit

Public Sub ApplySameFormatCharts()

Dim TemplateChartObject As ChartObject
Dim TemplateChartName As String
Dim TemplateChartWidth As Double
Dim TemplateChartHeight As Double
Dim TemplateChartType As XlChartType
Dim TemplateTitleForumla As String
Dim TemplateAxesFormulas As Collection

Dim Rtn_Size As Long
Dim Rtn_Temp As Long
Dim Rtn_Tite As Long
Dim Rtn_Axes As Long
Dim Rtn_Legd As Long
Dim Rtn_Shpe As Long

Dim TargetSheet As Worksheet
Dim TargetChartObject As ChartObject
Dim TargetChartAxes As Axes
Dim TargetItem As Collection
Dim TargetChartTitleData As Collection
Dim TargetChartLegendData As Collection
Dim TargetChartAxesData As Collection
Dim InsideShapesCount As Long
Dim ChangedInsideShapesCount As Long
Dim i As Long

    If ActiveChart Is Nothing Then
        MsgBox "グラフが選択されていません。グラフを選択してから実行してください。"
        Exit Sub
    End If
    
    Set TemplateChartObject = ActiveChart.Parent
    TemplateChartWidth = TemplateChartObject.Width
    TemplateChartHeight = TemplateChartObject.Height
    TemplateChartType = TemplateChartObject.Chart.ChartType
    If TemplateChartObject.Chart.HasTitle Then
        TemplateTitleForumla = TemplateChartObject.Chart.ChartTitle.Formula
    End If
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

    Rtn_Temp = MsgBox("テンプレートグラフの書式を適用しますか？", vbYesNo + vbQuestion)
    Rtn_Size = MsgBox("テンプレートグラフの大きさを適用しますか？", vbYesNo + vbQuestion)
    If Rtn_Temp = vbYes Then
        Rtn_Axes = MsgBox("テンプレートグラフの軸ラベル設定を適用しますか？", vbYesNo + vbQuestion)
        Rtn_Tite = MsgBox("テンプレートグラフのタイトル設定を適用しますか？", vbYesNo + vbQuestion)
        Rtn_Legd = MsgBox("テンプレートグラフの凡例設定を適用しますか？", vbYesNo + vbQuestion)
        Rtn_Shpe = MsgBox("テンプレートグラフの図形を追加しますか？", vbYesNo + vbQuestion)
    End If
    
    For Each TargetSheet In ActiveWorkbook.Sheets
        For Each TargetChartObject In TargetSheet.ChartObjects
            '対象グラフのタイトル数式・位置を取得
            Set TargetChartTitleData = New Collection
            If TargetChartObject.Chart.HasTitle Then
                TargetChartTitleData.Add TargetChartObject.Chart.ChartTitle.Formula, "Formula"
                TargetChartTitleData.Add TargetChartObject.Chart.ChartTitle.Top, "Top"
                TargetChartTitleData.Add TargetChartObject.Chart.ChartTitle.Left, "Left"
            End If
            '対象グラフの凡例の位置サイズ等を取得
            Set TargetChartLegendData = New Collection
            If TargetChartObject.Chart.HasLegend Then
                TargetChartLegendData.Add TargetChartObject.Chart.Legend.Position, "Position"
                TargetChartLegendData.Add TargetChartObject.Chart.Legend.Width, "Width"
                TargetChartLegendData.Add TargetChartObject.Chart.Legend.Height, "Height"
                TargetChartLegendData.Add TargetChartObject.Chart.Legend.Top, "Top"
                TargetChartLegendData.Add TargetChartObject.Chart.Legend.Left, "Left"
                TargetChartLegendData.Add TargetChartObject.Chart.Legend.Border.LineStyle, "LineStyle"
                If TargetChartObject.Chart.Legend.Border.LineStyle <> xlLineStyleNone Then
                    TargetChartLegendData.Add TargetChartObject.Chart.Legend.Border.Weight, "LineWeight"
                    If TargetChartObject.Chart.Legend.Border.ColorIndex = xlNone Then
                        TargetChartLegendData.Add True, "BorderColorIndex"
                    Else
                        TargetChartLegendData.Add False, "BorderColorIndex"
                        TargetChartLegendData.Add TargetChartObject.Chart.Legend.Border.Color, "BorderColor"
                    End If
                End If
                If TargetChartObject.Chart.Legend.Interior.ColorIndex = xlNone Then
                    TargetChartLegendData.Add True, "InteriorColorIndex"
                Else
                    TargetChartLegendData.Add False, "InteriorColorIndex"
                    TargetChartLegendData.Add TargetChartObject.Chart.Legend.Interior.Color, "InteriorColor"
                End If
            End If
            '対象グラフの軸ラベル数式・位置を取得
            Set TargetChartAxes = TargetChartObject.Chart.Axes
            Set TargetChartAxesData = New Collection
            i = 1
            Do Until i > TargetChartAxes.Count
                Set TargetItem = New Collection
                If TargetChartAxes.Item(i).HasTitle Then
                    TargetItem.Add TargetChartAxes.Item(i).AxisTitle.Formula, "Item" & i
                    TargetItem.Add TargetChartAxes.Item(i).AxisTitle.Top, "Top"
                    TargetItem.Add TargetChartAxes.Item(i).AxisTitle.Left, "Left"
                    TargetItem.Add TargetChartAxes.Item(i).AxisTitle.Orientation, "Orientation"
                Else
                    TargetItem.Add "なし", "Item" & i
                End If
                TargetChartAxesData.Add TargetItem
                i = i + 1
            Loop

            If TargetChartObject.Chart.ChartType = TemplateChartType Then
                If Rtn_Temp = vbYes Then
                    
                    InsideShapesCount = TargetChartObject.Chart.Shapes.Count
                    TargetChartObject.Chart.ApplyChartTemplate (TemplateChartName)
                    ChangedInsideShapesCount = TargetChartObject.Chart.Shapes.Count
                    
                    'グラフタイトル数式再設定
                    If Rtn_Tite = vbYes Then
                        If TemplateTitleForumla <> "" Then
                            TargetChartObject.Chart.ChartTitle.Formula = TemplateTitleForumla
                        Else
                            TargetChartObject.Chart.SetElement (msoElementChartTitleNone)
                        End If
                    Else
                        If TargetChartTitleData.Count > 0 Then
                            If TargetChartObject.Chart.HasTitle = False Then TargetChartObject.Chart.SetElement (msoElementChartTitleAboveChart)
                            TargetChartObject.Chart.ChartTitle.Formula = TargetChartTitleData("Formula")
                            TargetChartObject.Chart.ChartTitle.Top = TargetChartTitleData("Top")
                            TargetChartObject.Chart.ChartTitle.Left = TargetChartTitleData("Left")
                        Else
                            TargetChartObject.Chart.SetElement (msoElementChartTitleNone)
                        End If
                    End If
                    
                    '軸ラベル数式再設定
                    i = 1
                    Set TargetChartAxes = TargetChartObject.Chart.Axes
                    Do Until i > TargetChartAxes.Count
                       If Rtn_Axes = vbYes Then
                            If TemplateAxesFormulas("Item" & i) <> "なし" Then
                                TargetChartAxes.Item(i).AxisTitle.Formula = TemplateAxesFormulas("Item" & i)
                            Else
                                Select Case TargetChartAxes.Item(i).Type
                                    Case xlValue
                                        TargetChartObject.Chart.SetElement (msoElementPrimaryValueAxisTitleNone)
                                    Case xlCategory
                                        TargetChartObject.Chart.SetElement (msoElementPrimaryCategoryAxisTitleNone)
                                End Select
                            End If
                       Else
                            If TargetChartAxesData(i)("Item" & i) <> "なし" Then
                                If TargetChartAxes.Item(i).HasTitle = False Then
                                    Select Case TargetChartAxes.Item(i).Type
                                        Case xlValue
                                            TargetChartObject.Chart.SetElement (msoElementPrimaryValueAxisTitleBelowAxis)
                                        Case xlCategory
                                            TargetChartObject.Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                                    End Select
                                End If
                                TargetChartAxes.Item(i).AxisTitle.Formula = TargetChartAxesData(i)("Item" & i)
                                TargetChartAxes.Item(i).AxisTitle.Top = TargetChartAxesData(i)("Top")
                                TargetChartAxes.Item(i).AxisTitle.Left = TargetChartAxesData(i)("Left")
                                TargetChartAxes.Item(i).AxisTitle.Orientation = TargetChartAxesData(i)("Orientation")
                            Else
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
                    
                    '凡例を更新しない場合、凡例の設定を戻す
                    If Rtn_Legd = vbNo Then
                        If TargetChartLegendData.Count > 0 Then
                            If TargetChartObject.Chart.HasLegend = False Then TargetChartObject.Chart.SetElement (msoElementLegendLeftOverlay)

                            If TargetChartLegendData("Position") = xlLegendPositionCustom Then
                                TargetChartObject.Chart.Legend.Width = TargetChartLegendData("Width")
                                TargetChartObject.Chart.Legend.Height = TargetChartLegendData("Height")
                                TargetChartObject.Chart.Legend.Top = TargetChartLegendData("Top")
                                TargetChartObject.Chart.Legend.Left = TargetChartLegendData("Left")
                            Else
                                Select Case TargetChartLegendData("Position")
                                    Case xlLegendPositionBottom
                                        TargetChartObject.Chart.SetElement (msoElementLegendBottom)
                                    Case xlLegendPositionLeft
                                        TargetChartObject.Chart.SetElement (msoElementLegendLeft)
                                    Case xlLegendPositionRight
                                        TargetChartObject.Chart.SetElement (msoElementLegendRight)
                                    Case xlLegendPositionTop
                                        TargetChartObject.Chart.SetElement (msoElementLegendTop)
                                End Select
                            End If
                            TargetChartObject.Chart.Legend.Border.LineStyle = TargetChartLegendData("LineStyle")
                            If TargetChartLegendData("LineStyle") <> xlLineStyleNone Then
                                TargetChartObject.Chart.Legend.Border.Weight = TargetChartLegendData("LineWeight")
                                If TargetChartLegendData("BorderColorIndex") Then
                                    TargetChartObject.Chart.Legend.Border.ColorIndex = xlNone
                                Else
                                    TargetChartObject.Chart.Legend.Border.Color = TargetChartLegendData("BorderColor")
                                End If
                            End If
                            If TargetChartLegendData("InteriorColorIndex") Then
                                TargetChartObject.Chart.Legend.Interior.ColorIndex = xlNone
                            Else
                                TargetChartObject.Chart.Legend.Interior.Color = TargetChartLegendData("InteriorColor")
                            End If
                        Else
                            TargetChartObject.Chart.SetElement (msoElementLegendNone)
                        End If
                    End If
                    
                    '追加しない場合、追加された図形を削除
                    If Rtn_Shpe = vbNo Then
                        If InsideShapesCount <> ChangedInsideShapesCount Then
                            For i = ChangedInsideShapesCount To InsideShapesCount + 1 Step -1
                                TargetChartObject.Chart.Shapes(i).Delete
                            Next i
                        End If
                    End If
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

