Attribute VB_Name = "Module1"
Option Explicit

Private TargetBook As Workbook
Private TargetFileName As String
Private OutputSheet As Worksheet

Public Sub FindExternalLinks()

Dim i As Long
Dim myLinkSources As Variant
Dim Fso  As Object
Dim myDic As Object

    Application.ScreenUpdating = False
    
    Set OutputSheet = ThisWorkbook.Sheets("出力")
    
    OutputSheet.Range("A1").CurrentRegion.Offset(1).ClearContents
    
    Set TargetBook = Workbooks.Open(Filename:=ThisWorkbook.Sheets("設定").Range("B2").Value, UpdateLinks:=False)
    
    '非表示シートを表示
    For i = 1 To TargetBook.Worksheets.Count
    
        TargetBook.Worksheets(i).Visible = xlSheetVisible
        
    Next i

    'ブックリンク情報の取得
    myLinkSources = TargetBook.LinkSources
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set myDic = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(myLinkSources)
    
        If Not myDic.Exists(Fso.GetFileName(myLinkSources(i))) Then
        
            myDic.Add Fso.GetFileName(myLinkSources(i)), Fso.GetFileName(myLinkSources(i))
            
        End If
        
    Next i
    
    Set Fso = Nothing
    
    For i = 0 To myDic.Count - 1
    
        TargetFileName = myDic.Items()(i)
      
        Call SearchCells
        Call SearchNames
        Call SearchValidation
        Call SearchFormatConditions
        Call SearchShapes

    Next i
    
    ThisWorkbook.Activate
    OutputSheet.Activate
    
    '画面左上に移動
    Application.Goto Reference:=OutputSheet.Range("A1"), Scroll:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "外部参照ブックリンク検索が完了しました。"

    
End Sub

Private Sub SearchCells()
    
Dim FindRange As Range
Dim mySheet As Worksheet
Dim TargetRow As Long
Dim StartFindRange As Range
    
    For Each mySheet In TargetBook.Worksheets
          
        Set FindRange = mySheet.Cells.Find(TargetFileName, LookIn:=xlFormulas, Lookat:=xlPart)
        
        If Not FindRange Is Nothing Then
            
            Set StartFindRange = FindRange
                 
            If FindRange.HasFormula = True And InStr(FindRange.Formula, "[") > 0 Then
            
                TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                OutputSheet.Cells(TargetRow, "C").Value = "セル"
                OutputSheet.Cells(TargetRow, "D").Value = FindRange.Address
                OutputSheet.Cells(TargetRow, "E").Value = "'" & FindRange.Formula
                
            End If
            
            '次の検索
            Do
            
                Set FindRange = mySheet.Cells.FindNext(FindRange)
                
                '最初に見つかったセルに戻ったら終了
                If StartFindRange.Address = FindRange.Address Then Exit Do
    
                If FindRange.HasFormula And InStr(FindRange.Formula, "[") > 0 Then
                 
                     TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                     OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                     OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                     OutputSheet.Cells(TargetRow, "C").Value = "セル"
                     OutputSheet.Cells(TargetRow, "D").Value = FindRange.Address
                     OutputSheet.Cells(TargetRow, "E").Value = "'" & FindRange.Formula
                 
                 End If
            
            Loop
            
        End If
  
    Next mySheet
 
End Sub

Private Sub SearchNames()

Dim TargetRow As Long
Dim MyName As Name
Dim i As Long


    '非表示の名前定義を表示
    For i = 1 To TargetBook.Names.Count
    
        TargetBook.Names.Item(i).Visible = True
    
    Next i

    For Each MyName In TargetBook.Names

        If MyName.Value Like "*" & TargetFileName & "*" Then
            
             TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
             OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
             OutputSheet.Cells(TargetRow, "B").Value = "―"
             OutputSheet.Cells(TargetRow, "C").Value = "名前定義"
             OutputSheet.Cells(TargetRow, "D").Value = MyName.Name
             OutputSheet.Cells(TargetRow, "E").Value = "'" & MyName.Value
             
        End If
      
    Next MyName
        
End Sub

Private Sub SearchValidation()

Dim mySheet As Worksheet
Dim myRange As Range
Dim mySameRange As Range
Dim myValidation As Validation
Dim iCount As Long
Dim TargetRow As Long
Dim myDic As Object


   Set myDic = CreateObject("Scripting.Dictionary")
   
   On Error Resume Next
          
   For Each mySheet In TargetBook.Worksheets
         
        iCount = 0
         
        'エラーの場合（＝入力規則が存在しない）　値は0のままで次へ
        iCount = mySheet.Cells.SpecialCells(xlCellTypeAllValidation).Count

        If iCount <> 0 Then
            
            '対象セルごと
            For Each myRange In mySheet.Cells.SpecialCells(xlCellTypeAllValidation)
                
                '同じ入力規則が設定されているセルをまとめる
                Set mySameRange = myRange.SpecialCells(xlCellTypeSameValidation)
                
                'Dictionaryにまとめた範囲がない場合、Dictionaryに追加
                If Not myDic.Exists(mySameRange.Address) Then
        
                    myDic.Add mySameRange.Address, mySameRange.Address
        
                    Set myValidation = mySameRange.Validation

                    If myValidation.Formula1 Like "*" & TargetFileName & "*" Then
                        
                        TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                        OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                        OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                        OutputSheet.Cells(TargetRow, "C").Value = "入力規則 "
                        OutputSheet.Cells(TargetRow, "D").Value = mySameRange.Address
                        OutputSheet.Cells(TargetRow, "E").Value = "'" & myValidation.Formula1
  
                    ElseIf myValidation.Operator = xlBetween Or myValidation.Operator = xlNotBetween Then
                    
                        If myValidation.Formula2 Like "*" & TargetFileName & "*" Then
                            
                            TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                            OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                            OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                            OutputSheet.Cells(TargetRow, "C").Value = "入力規則 "
                            OutputSheet.Cells(TargetRow, "D").Value = mySameRange.Address
                            OutputSheet.Cells(TargetRow, "E").Value = "'" & myValidation.Formula2
                        
                        End If

                    End If
                    
                End If
                    
            Next myRange
        
        End If
            
    Next mySheet
    
    On Error GoTo 0
    
End Sub

Private Sub SearchFormatConditions()

Dim mySheet As Worksheet
Dim myFormatCondition As Object
Dim TargetRow As Long
Dim TargetObj As Object

    For Each mySheet In TargetBook.Worksheets
        
        For Each myFormatCondition In mySheet.Cells.FormatConditions
                           
            Select Case myFormatCondition.Type
            
                Case xlCellValue, xlTextString, xlExpression
                    
                    If myFormatCondition.Formula1 Like "*" & TargetFileName & "*" Then
                    
                        TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                        OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                        OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                        OutputSheet.Cells(TargetRow, "C").Value = "条件付き書式"
                        OutputSheet.Cells(TargetRow, "D").Value = myFormatCondition.AppliesTo.Address
                        OutputSheet.Cells(TargetRow, "E").Value = "'" & myFormatCondition.Formula1

                    ElseIf myFormatCondition.Type = xlCellValue Then
                        
                        If myFormatCondition.Operator = xlBetween Or myFormatCondition.Operator = xlNotBetween Then
                            
                            If myFormatCondition.Formula2 Like "*" & TargetFileName & "*" Then
                            
                                TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                                OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                                OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                                OutputSheet.Cells(TargetRow, "C").Value = "条件付き書式 セルの値"
                                OutputSheet.Cells(TargetRow, "D").Value = myFormatCondition.AppliesTo.Address
                                OutputSheet.Cells(TargetRow, "E").Value = "'" & myFormatCondition.Formula2
                                
                            End If
                        
                        End If
                        
                    End If

                Case xlColorScale
                    
                    For Each TargetObj In myFormatCondition.ColorScaleCriteria
                        
                        If TargetObj.Value Like "*" & TargetFileName & "*" Then
                            
                            TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                            OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                            OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                            OutputSheet.Cells(TargetRow, "C").Value = "条件付き書式"
                            OutputSheet.Cells(TargetRow, "D").Value = myFormatCondition.AppliesTo.Address
                            OutputSheet.Cells(TargetRow, "E").Value = "'" & TargetObj.Value
                            Exit For
                            
                        End If
                        
                    Next TargetObj
                    
                    
                Case xlIconSets
                    
                    For Each TargetObj In myFormatCondition.IconCriteria
                       
                       If TargetObj.Value Like "*" & TargetFileName & "*" Then
                            
                            TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                            OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                            OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                            OutputSheet.Cells(TargetRow, "C").Value = "条件付き書式"
                            OutputSheet.Cells(TargetRow, "D").Value = myFormatCondition.AppliesTo.Address
                            OutputSheet.Cells(TargetRow, "E").Value = "'" & TargetObj.Value
                            Exit For
                                
                        End If
                        
                    Next TargetObj
                
                
                Case xlDatabar
                
                    If myFormatCondition.MaxPoint.Value Like "*" & TargetFileName & "*" Then
                        
                        TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                        OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                        OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                        OutputSheet.Cells(TargetRow, "C").Value = "条件付き書式"
                        OutputSheet.Cells(TargetRow, "D").Value = myFormatCondition.AppliesTo.Address
                        OutputSheet.Cells(TargetRow, "E").Value = "'" & myFormatCondition.MaxPoint.Value
                        
                    ElseIf myFormatCondition.MinPoint.Value Like "*" & TargetFileName & "*" Then
                        
                        TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                        OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                        OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                        OutputSheet.Cells(TargetRow, "C").Value = "条件付き書式"
                        OutputSheet.Cells(TargetRow, "D").Value = myFormatCondition.AppliesTo.Address
                        OutputSheet.Cells(TargetRow, "E").Value = "'" & myFormatCondition.MaxPoint.Value
                        
                    End If

            End Select
        
        Next myFormatCondition
        
    Next mySheet

End Sub

Private Sub SearchShapes()

Dim mySheet As Worksheet
Dim myShape As Shape
Dim TargetRow As Long
 
    For Each mySheet In TargetBook.Worksheets
    
        For Each myShape In mySheet.Shapes
            
            Call SerchShapeProcess(mySheet, myShape)
        
        Next myShape

    Next mySheet
    
End Sub



Private Sub SerchShapeProcess(mySheet As Worksheet, myShape As Shape)

Dim TargetRow As Long
Dim myShapeTypeName As String
Dim myGroupShape As Shape
Dim TargetAxis As Object
Dim mySeries As Series
Dim ChackSheet As Worksheet
Dim ChackFormulaString As String
Dim ChackOnActionString As String

    myShapeTypeName = "図形"
    
    Select Case myShape.Type
        
        Case msoGroup
            
            For Each myGroupShape In myShape.GroupItems
            
                Call SerchShapeProcess(mySheet, myGroupShape)
            
            Next myGroupShape
     
        Case msoChart
            
            myShapeTypeName = "グラフ"
            
            If myShape.Chart.HasTitle Then
                
                If myShape.Chart.ChartTitle.Formula Like "*" & TargetFileName & "*" Then
                         
                    TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                    OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                    OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                    OutputSheet.Cells(TargetRow, "C").Value = "グラフタイトル"
                    OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                    OutputSheet.Cells(TargetRow, "E").Value = "'" & myShape.Chart.ChartTitle.Formula
                    
                End If
                
            End If

             
            For Each TargetAxis In myShape.Chart.Axes
                
                If TargetAxis.HasTitle Then
                    
                    If TargetAxis.AxisTitle.Formula Like "*" & TargetFileName & "*" Then

                         TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                         OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                         OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                         
                         Select Case TargetAxis.Type

                            Case xlValue
                            
                                OutputSheet.Cells(TargetRow, "C").Value = "グラフ数値軸タイトル"
                         
                            Case xlCategory
                                
                                OutputSheet.Cells(TargetRow, "C").Value = "グラフ項目軸タイトル"
                            
                            Case Else
                            
                                OutputSheet.Cells(TargetRow, "C").Value = "グラフ軸タイトル"
                                
                         End Select
                         
                         OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                         OutputSheet.Cells(TargetRow, "E").Value = "'" & TargetAxis.AxisTitle.Formula
                    
                    End If

                End If

            Next TargetAxis

            For Each mySeries In myShape.Chart.FullSeriesCollection
        
                If mySeries.Formula Like "*" & TargetFileName & "*" Then
                
                    TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                    OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                    OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                    OutputSheet.Cells(TargetRow, "C").Value = "グラフ系列範囲 / " & mySeries.Name
                    OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                    OutputSheet.Cells(TargetRow, "E").Value = "'" & mySeries.Formula

                End If
             
            Next mySeries

        Case msoFormControl
            
            '入力規則のドロップダウンの確認
            If myShape.FormControlType = xlDropDown Then
                
                If Not DropDownChack(mySheet, myShape) Then Exit Sub
            
            End If
            
            myShapeTypeName = "フォームコントロール"
            
            Select Case myShape.FormControlType
                
                Case xlLabel, xlGroupBox
            
                    Set ChackSheet = TargetBook.Sheets.Add
                
                     myShape.Copy
                     ChackSheet.Paste
                     
                     If ChackSheet.DrawingObjects.LinkedCell Like "*" & TargetFileName & "*" Then
                                  
                         TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                         
                         OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                         OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                         OutputSheet.Cells(TargetRow, "C").Value = "フォームコントロール / セル参照"
                         OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                         OutputSheet.Cells(TargetRow, "E").Value = "'" & ChackSheet.DrawingObjects.LinkedCell
                         
                     End If
                     
                     Application.DisplayAlerts = False
                    
                     ChackSheet.Delete
                    
                     Application.DisplayAlerts = True
                     
                
                Case Is <> xlButtonControl
                            
                    If myShape.DrawingObject.LinkedCell Like "*" & TargetFileName & "*" Then
                              
                        TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                        OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                        OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                        OutputSheet.Cells(TargetRow, "C").Value = "フォームコントロール / リンクするセル"
                        OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                        OutputSheet.Cells(TargetRow, "E").Value = "'" & myShape.DrawingObject.LinkedCell

                    End If
                    
                    If myShape.FormControlType = xlListBox Or myShape.FormControlType = xlDropDown Then
            
                        If myShape.DrawingObject.ListFillRange Like "*" & TargetFileName & "*" Then

                            TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                            OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                            OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                            OutputSheet.Cells(TargetRow, "C").Value = "フォームコントロール / 入力範囲"
                            OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                            OutputSheet.Cells(TargetRow, "E").Value = "'" & myShape.DrawingObject.ListFillRange

                        End If
                        
                    End If

            End Select
    
    End Select
     
    If myShape.Type <> msoGroup Then
        
        On Error Resume Next
        
        'Formulaプロパティがない場合、エラーで空白のまま次へ
        ChackFormulaString = myShape.DrawingObject.Formula
        
        If ChackFormulaString <> "" Then

            If myShape.DrawingObject.Formula Like "*" & TargetFileName & "*" Then
 
                TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                OutputSheet.Cells(TargetRow, "C").Value = myShapeTypeName & " / セル参照"
                OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                OutputSheet.Cells(TargetRow, "E").Value = myShape.DrawingObject.Formula
    
            End If
        
        End If
   
        'OnActionプロパティがない場合、エラーで空白のまま次へ
        ChackOnActionString = myShape.OnAction
         
        If ChackOnActionString <> "" Then

            If myShape.OnAction Like "*" & TargetFileName & "*" Then
   
                TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
                OutputSheet.Cells(TargetRow, "B").Value = mySheet.Name
                OutputSheet.Cells(TargetRow, "C").Value = myShapeTypeName & " / マクロ登録"
                OutputSheet.Cells(TargetRow, "D").Value = myShape.Name
                OutputSheet.Cells(TargetRow, "E").Value = myShape.OnAction
 
            End If
        
        End If
        
    End If
    
    On Error GoTo 0
  
End Sub


'入力規則のドロップダウン除外用
Private Function DropDownChack(mySheet As Worksheet, myShape As Shape) As Boolean
Dim myDropDown As DropDown

    DropDownChack = False
    
    For Each myDropDown In mySheet.DropDowns
        
        If myDropDown.Name = myShape.Name Then DropDownChack = True: Exit For
      
    Next myDropDown

End Function

