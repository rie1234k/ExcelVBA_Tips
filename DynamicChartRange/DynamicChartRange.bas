Attribute VB_Name = "Module1"
Option Explicit


Public Sub Set_DynamicChartRange()

Dim TargetWorksheet As Worksheet
Dim StartItemRange  As Range
Dim EndItemRange As Range

Dim TargetSeriesCollection As Collection
Dim TargetChartObject As ChartObject
Dim TargetSeries As Series
Dim myItem As Collection
Dim TargetDirection As String

Dim StartFormula As String
Dim EndFormula As String
Dim CountFormula As String
Dim TargetAxesStr As String
Dim TargetAxesStartStr As String
Dim TargetSeriesStr As String
Dim TargetSeriesStartStr As String
Dim TargetSeriesIndex As String
Dim CharStart As Long
Dim CharLength As Long
Dim ChangeFormula As String
  
    '------- 設定開始 -------
    Set TargetWorksheet = ActiveSheet
    Set StartItemRange = TargetWorksheet.Range("C3") '開始項目入力セルを設定
    Set EndItemRange = TargetWorksheet.Range("C4") '終了項目入力セルを設定
    '------- 設定終了 -------

     TargetWorksheet.Names.Add Name:=TargetWorksheet.Name & "_範囲開始", RefersTo:="='" & TargetWorksheet.Name & "'!" & StartItemRange.Address
     TargetWorksheet.Names.Add Name:=TargetWorksheet.Name & "_範囲終了", RefersTo:="='" & TargetWorksheet.Name & "'!" & EndItemRange.Address
     
     Set TargetSeriesCollection = New Collection
                       
     For Each TargetChartObject In TargetWorksheet.ChartObjects
         
         For Each TargetSeries In TargetChartObject.Chart.FullSeriesCollection
            
            CharStart = InStr(TargetSeries.Formula, ",") + 1
            CharLength = InStr(CharStart, TargetSeries.Formula, ",") - CharStart
            TargetAxesStr = Mid(TargetSeries.Formula, CharStart, CharLength)
            
            CharStart = InStr(CharStart, TargetSeries.Formula, ",") + 1
            CharLength = InStr(CharStart, TargetSeries.Formula, ",") - CharStart
            TargetSeriesStr = Mid(TargetSeries.Formula, CharStart, CharLength)
            TargetSeriesIndex = Replace(Mid(TargetSeries.Formula, InStrRev(TargetSeries.Formula, ",") + 1), ")", "")
             
            '連続するデータのみ対象
            If Len(TargetSeries.Formula) - Len(Replace(TargetSeries.Formula, ",", "")) = 3 And InStr(TargetSeriesStr, ":") > 0 Then

                '開始・終了項目がラベル範囲にあるか
                If WorksheetFunction.CountIf(TargetWorksheet.Range(TargetAxesStr), StartItemRange.Value) _
                    And WorksheetFunction.CountIf(TargetWorksheet.Range(TargetAxesStr), EndItemRange.Value) Then
                    
                    'ラベルの方向確認
                    If TargetWorksheet.Range(TargetAxesStr).Columns.Count > TargetWorksheet.Range(TargetAxesStr).Rows.Count Then
                        
                        TargetDirection = "横"
                        
                    Else
                    
                        TargetDirection = "縦"
                        
                    End If

                    Set myItem = New Collection
                    
                    myItem.Add TargetSeries, "系列"
                    myItem.Add TargetAxesStr, "軸ラベル範囲"
                    myItem.Add TargetWorksheet.Range(TargetAxesStr).Address(ReferenceStyle:=xlR1C1), "軸ラベル範囲R1C1"
                    myItem.Add TargetDirection, "軸ラベル方向"
                    myItem.Add TargetSeriesStr, "系列範囲"
                    myItem.Add TargetWorksheet.Range(TargetSeriesStr).Address(ReferenceStyle:=xlR1C1), "系列範囲R1C1"
                    myItem.Add "系列範囲" & TargetSeriesIndex, "系列名"
                    myItem.Add Replace(TargetChartObject.Name, " ", ""), "グラフ名"
                    
                    TargetSeriesCollection.Add myItem
                    
                    Set myItem = Nothing
                
                End If
                
            End If
            
         Next TargetSeries
     
     Next TargetChartObject
 

     For Each myItem In TargetSeriesCollection
            
        'ラベル範囲の起点セルアドレス
        TargetAxesStartStr = Left(myItem("軸ラベル範囲"), InStr(myItem("軸ラベル範囲"), ":") - 1)
        
        'ラベル(行・列)全体の名前定義
        Select Case myItem("軸ラベル方向")
        
            Case "横"
            
                TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_軸ラベル範囲全体", RefersTo:="=" & TargetWorksheet.Range(TargetAxesStartStr).EntireRow.Address
           
            Case "縦"
                
                TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_軸ラベル範囲全体", RefersTo:="=" & TargetWorksheet.Range(TargetAxesStartStr).EntireColumn.Address
        
        End Select
        
        '開始位置、表示件数の名前定義
        StartFormula = "MATCH(" & TargetWorksheet.Name & "_範囲開始," & myItem("グラフ名") & "_軸ラベル範囲全体" & ",0)"
        EndFormula = "MATCH(" & TargetWorksheet.Name & "_範囲終了," & myItem("グラフ名") & "_軸ラベル範囲全体" & ",0)"
        CountFormula = EndFormula & " - " & StartFormula & " +1"

        TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_開始位置", RefersTo:="=" & StartFormula
        TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_表示件数", RefersTo:="=" & CountFormula
        
         '系列範囲の起点セルアドレス
        TargetSeriesStartStr = Left(myItem("系列範囲"), InStr(myItem("系列範囲"), ":") - 1)
        
        '指定軸ラベル範囲、系列の名前定義
        Select Case myItem("軸ラベル方向")
              
              Case "横"
                
                TargetAxesStartStr = Replace(TargetAxesStartStr, Mid(TargetAxesStartStr, InStr(TargetAxesStartStr, "$") + 1), "A") & Mid(TargetAxesStartStr, InStrRev(TargetAxesStartStr, "$"))
                TargetSeriesStartStr = Replace(TargetSeriesStartStr, Mid(TargetSeriesStartStr, InStr(TargetAxesStartStr, "$") + 1), "A") & Mid(TargetSeriesStartStr, InStrRev(TargetSeriesStartStr, "$"))
                TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_指定軸ラベル範囲", RefersTo:="=OFFSET(" & TargetAxesStartStr & ",0," & myItem("グラフ名") & "_開始位置 -1,1," & myItem("グラフ名") & "_表示件数)"
                TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_" & myItem("系列名"), RefersTo:="=OFFSET(" & TargetSeriesStartStr & ",0," & myItem("グラフ名") & "_開始位置 -1,1," & myItem("グラフ名") & "_表示件数)"

            Case "縦"
            
                TargetAxesStartStr = Replace(TargetAxesStartStr, Mid(TargetAxesStartStr, InStrRev(TargetAxesStartStr, "$") + 1), "1")
                TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_指定軸ラベル範囲", RefersTo:="=OFFSET(" & TargetAxesStartStr & "," & myItem("グラフ名") & "_開始位置 -1,0," & myItem("グラフ名") & "_表示件数,1)"
                TargetSeriesStartStr = Replace(TargetSeriesStartStr, Mid(TargetSeriesStartStr, InStrRev(TargetSeriesStartStr, "$") + 1), "1")
                TargetWorksheet.Names.Add Name:=myItem("グラフ名") & "_" & myItem("系列名"), RefersTo:="=OFFSET(" & TargetSeriesStartStr & "," & myItem("グラフ名") & "_開始位置 -1,0," & myItem("グラフ名") & "_表示件数,1)"

        End Select
        
        ChangeFormula = Replace(myItem("系列").FormulaR1C1, myItem("軸ラベル範囲R1C1"), myItem("グラフ名") & "_指定軸ラベル範囲")
        ChangeFormula = Replace(ChangeFormula, myItem("系列範囲R1C1"), myItem("グラフ名") & "_" & myItem("系列名"))
        
        myItem("系列").FormulaR1C1 = ChangeFormula
        
    Next myItem
     
    MsgBox "グラフ設定が完了しました"
    
     
End Sub






