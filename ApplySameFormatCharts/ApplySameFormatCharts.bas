Attribute VB_Name = "Module1"
Option Explicit


Public Sub ApplySameFormatCharts()

Dim TemplateChartObject As ChartObject
Dim TemplateChartName As String
Dim TemplateChartWidth As Double
Dim TemplateChartHeight As Double
Dim TemplateChartType As XlChartType
  
Dim Rtn_Size As Long
Dim Rtn_Temp As Long

Dim TargetSheet As Worksheet
Dim TargetChartObject As ChartObject


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
    
    
    Rtn_Temp = MsgBox("選択中のグラフの書式を適用しますか？", vbYesNo + vbQuestion)
    Rtn_Size = MsgBox("選択中のグラフの大きさを適用しますか？", vbYesNo + vbQuestion)
     
    
     For Each TargetSheet In ActiveWorkbook.Sheets
      
        For Each TargetChartObject In TargetSheet.ChartObjects
        
            If TargetChartObject.Chart.ChartType = TemplateChartType Then
                
                If Rtn_Temp = vbYes Then
                    
                    TargetChartObject.Chart.ApplyChartTemplate (TemplateChartName)
    
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
