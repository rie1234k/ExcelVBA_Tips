Attribute VB_Name = "Module1"
Option Explicit

Private TargetBook As Workbook
Private TargetFileName As String
Private OutputSheet As Worksheet
Private ChackSheet As Worksheet


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
    
    MsgBox "外部のブックリンク検索が完了しました。"

    
End Sub

Private Sub SearchCells()
    
Dim FindRange As Range
Dim mySheet As Worksheet
Dim TargetRow As Long
Dim StartFindRange As Range
    
    For Each mySheet In TargetBook.Worksheets
        
        'オートフィルタクリア
        If mySheet.FilterMode Then ActiveSheet.ShowAllData
        
        '非表示行・列を表示
        mySheet.Cells.EntireRow.Hidden = False
        mySheet.Cells.EntireColumn.Hidden = False
          
        Set FindRange = mySheet.Cells.Find(TargetFileName, LookIn:=xlFormulas, Lookat:=xlPart)
        
        If Not FindRange Is Nothing Then
            
            Set StartFindRange = FindRange
                 
            If FindRange.HasFormula = True And InStr(FindRange.Formula, "[") > 0 Then
                
                Call OutputProcess(mySheet.Name, "セル", FindRange.Address, "'" & FindRange.Formula)

            End If
            
            '次の検索
            Do
            
                Set FindRange = mySheet.Cells.FindNext(FindRange)
                
                '最初に見つかったセルに戻ったら終了
                If StartFindRange.Address = FindRange.Address Then Exit Do
    
                If FindRange.HasFormula And InStr(FindRange.Formula, "[") > 0 Then
                    
                    Call OutputProcess(mySheet.Name, "セル", FindRange.Address, "'" & FindRange.Formula)
                 
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
            
            Call OutputProcess("―", "名前定義", MyName.Name, "'" & MyName.Value)
             
        End If
      
    Next MyName
        
End Sub

Private Sub SearchValidation()

Dim mySheet As Worksheet
Dim myRange As Range
Dim mySameRange As Range
Dim myValidation As Validation
Dim iCount As Long
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
        
                    Set myValidation = mySameRange.Cells(1).Validation

                    If myValidation.Formula1 Like "*" & TargetFileName & "*" Then
                        
                        Call OutputProcess(mySheet.Name, "入力規則", mySameRange.Address, "'" & myValidation.Formula1)
  
                    ElseIf myValidation.Operator = xlBetween Or myValidation.Operator = xlNotBetween Then
                    
                        If myValidation.Formula2 Like "*" & TargetFileName & "*" Then
                            
                            Call OutputProcess(mySheet.Name, "入力規則", mySameRange.Address, "'" & myValidation.Formula2)
                                                
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
Dim TargetObj As Object

    For Each mySheet In TargetBook.Worksheets
        
        For Each myFormatCondition In mySheet.Cells.FormatConditions
                           
            Select Case myFormatCondition.Type
            
                Case xlCellValue, xlTextString, xlExpression
                    
                    If myFormatCondition.Formula1 Like "*" & TargetFileName & "*" Then
                        
                        Call OutputProcess(mySheet.Name, "条件付き書式", myFormatCondition.AppliesTo.Address, "'" & myFormatCondition.Formula1)
                        
                    ElseIf myFormatCondition.Type = xlCellValue Then
                        
                        If myFormatCondition.Operator = xlBetween Or myFormatCondition.Operator = xlNotBetween Then
                            
                            If myFormatCondition.Formula2 Like "*" & TargetFileName & "*" Then
                            
                                Call OutputProcess(mySheet.Name, "条件付き書式", myFormatCondition.AppliesTo.Address, "'" & myFormatCondition.Formula2)
                                
                            End If
                        
                        End If
                        
                    End If

                Case xlColorScale
                    
                    For Each TargetObj In myFormatCondition.ColorScaleCriteria
                        
                        If TargetObj.Value Like "*" & TargetFileName & "*" Then
                            
                            Call OutputProcess(mySheet.Name, "条件付き書式", myFormatCondition.AppliesTo.Address, "'" & TargetObj.Value)
                            Exit For
                            
                        End If
                        
                    Next TargetObj
                    
                    
                Case xlIconSets
                    
                    For Each TargetObj In myFormatCondition.IconCriteria
                       
                       If TargetObj.Value Like "*" & TargetFileName & "*" Then
                            
                            Call OutputProcess(mySheet.Name, "条件付き書式", myFormatCondition.AppliesTo.Address, "'" & TargetObj.Value)
                            Exit For
                                
                        End If
                        
                    Next TargetObj
                
                
                Case xlDatabar
                
                    If myFormatCondition.MaxPoint.Value Like "*" & TargetFileName & "*" Then
                        
                        Call OutputProcess(mySheet.Name, "条件付き書式", myFormatCondition.AppliesTo.Address, "'" & myFormatCondition.MaxPoint.Value)
                        
                    ElseIf myFormatCondition.MinPoint.Value Like "*" & TargetFileName & "*" Then
                        
                        Call OutputProcess(mySheet.Name, "条件付き書式", myFormatCondition.AppliesTo.Address, "'" & myFormatCondition.MaxPoint.Value)
                        
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
        
            If myShape.Visible = msoFalse Then myShape.Visible = msoCTrue
            
            Call SerchShapeProcess(mySheet, myShape)
        
        Next myShape

    Next mySheet
    
    'チェック用シートを作成している場合削除
    If Not ChackSheet Is Nothing Then
    
        Application.DisplayAlerts = False
                    
        ChackSheet.Delete
        
        Application.DisplayAlerts = True
    
    End If
    
End Sub



Private Sub SerchShapeProcess(mySheet As Worksheet, myShape As Shape)

Dim myShapeTypeName As String
Dim myGroupShape As Shape
Dim TargetAxis As Object
Dim AxisName As String
Dim mySeries As Series
Dim ChackFormulaString As String
Dim ChackOnActionString As String

    myShapeTypeName = "図形"
    
    Select Case myShape.Type
        
        Case msoGroup
            
            For Each myGroupShape In myShape.GroupItems
                
                If myGroupShape.Visible = msoFalse Then myGroupShape.Visible = msoCTrue
                 
                Call SerchShapeProcess(mySheet, myGroupShape)
            
            Next myGroupShape
     
        Case msoChart
            
            myShapeTypeName = "グラフ"
            
            If myShape.Chart.HasTitle Then
                
                If myShape.Chart.ChartTitle.Formula Like "*" & TargetFileName & "*" Then
                         
                    Call OutputProcess(mySheet.Name, "グラフタイトル", myShape.Name, "'" & myShape.Chart.ChartTitle.Formula)
                    
                End If
                
            End If

             
            For Each TargetAxis In myShape.Chart.Axes
                
                If TargetAxis.HasTitle Then
                    
                    If TargetAxis.AxisTitle.Formula Like "*" & TargetFileName & "*" Then
                         
                         Select Case TargetAxis.Type

                            Case xlValue
                            
                                AxisName = "グラフ数値軸タイトル"
                         
                            Case xlCategory
                                
                                AxisName = "グラフ項目軸タイトル"
                            
                            Case Else
                            
                                AxisName = "グラフ軸タイトル"
                                
                         End Select
                         
                         Call OutputProcess(mySheet.Name, AxisName, myShape.Name, "'" & TargetAxis.AxisTitle.Formula)
                         
                    
                    End If

                End If

            Next TargetAxis

            For Each mySeries In myShape.Chart.FullSeriesCollection
        
                If mySeries.Formula Like "*" & TargetFileName & "*" Then
                    
                    Call OutputProcess(mySheet.Name, "グラフ系列範囲 / " & mySeries.Name, myShape.Name, "'" & mySeries.Formula)
                    
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
            
                    If ChackSheet Is Nothing Then Set ChackSheet = TargetBook.Sheets.Add
                
                    myShape.Copy
                    Application.Wait Now() + TimeSerial(0, 0, 1)
                    ChackSheet.Paste
                     
                    If ChackSheet.DrawingObjects.LinkedCell Like "*" & TargetFileName & "*" Then
                                  
                        Call OutputProcess(mySheet.Name, "フォームコントロール / セル参照", myShape.Name, "'" & ChackSheet.DrawingObjects.LinkedCell)
                         
                    End If
                     
                    ChackSheet.Shapes(1).Delete
                    
                     
                
                Case Is <> xlButtonControl
                            
                    If myShape.DrawingObject.LinkedCell Like "*" & TargetFileName & "*" Then
                    
                        Call OutputProcess(mySheet.Name, "フォームコントロール / リンクするセル", myShape.Name, "'" & myShape.DrawingObject.LinkedCell)

                    End If
                    
                    If myShape.FormControlType = xlListBox Or myShape.FormControlType = xlDropDown Then
            
                        If myShape.DrawingObject.ListFillRange Like "*" & TargetFileName & "*" Then
                        
                            Call OutputProcess(mySheet.Name, "フォームコントロール / 入力範囲", myShape.Name, "'" & "'" & myShape.DrawingObject.ListFillRange)

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
                
                Call OutputProcess(mySheet.Name, myShapeTypeName & " / セル参照", myShape.Name, "'" & "'" & myShape.DrawingObject.Formula)
    
            End If
        
        End If
   
        'OnActionプロパティがない場合、エラーで空白のまま次へ
        ChackOnActionString = myShape.OnAction
         
        If ChackOnActionString <> "" Then

            If myShape.OnAction Like "*" & TargetFileName & "*" Then
                
                Call OutputProcess(mySheet.Name, myShapeTypeName & " / マクロ登録", myShape.Name, "'" & "'" & myShape.OnAction)
 
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

Private Sub OutputProcess(FindSheetName As String, FindTypeName As String, FindPlaceName As String, FindDetail As String)

Dim TargetRow As Long

    TargetRow = OutputSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
    OutputSheet.Cells(TargetRow, "A").Value = TargetFileName
    OutputSheet.Cells(TargetRow, "B").Value = FindSheetName
    OutputSheet.Cells(TargetRow, "C").Value = FindTypeName
    OutputSheet.Cells(TargetRow, "D").Value = FindPlaceName
    OutputSheet.Cells(TargetRow, "E").Value = FindDetail


End Sub
