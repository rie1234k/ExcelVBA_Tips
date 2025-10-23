Attribute VB_Name = "A_予定転記"
Option Explicit


Public Sub Write_Calendar()

Dim i As Long
Dim j As Long
Dim k As Long

Dim endRow As Long
Dim TargetRange As Range
Dim RowCount As Long
Dim TargetDate As Date
Dim CalendarSheet As Worksheet
Dim myData As Collection
Dim myDataTable As Collection
Dim myTableCollection As Collection
Dim FindRange As Range
Dim TargetData As Collection
Dim TargetAddress As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    With ThisWorkbook.Sheets("予定一覧")
    
        '------- 予定一覧 並び替え -------
    
        endRow = .Cells(Rows.Count, "B").Row
        Set TargetRange = .Range(.Range("B2"), .Cells(endRow, "F"))
        
        If WorksheetFunction.CountA(.Columns("B")) > 3 Then
                .Sort.SortFields.Clear
                .Sort.SortFields.Add Key:=.Range("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .Sort.SortFields.Add Key:=.Range("C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .Sort.SetRange TargetRange
            With .Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
        
        '------- カレンダー行数調整用 日付件数取得 -------
        i = 3
        
        Set myDataTable = New Collection
        
        Do
            If IsDate(.Cells(i, "B").Value) Then
                
                If .Cells(i, "B").Value <> .Cells(i - 1, "B").Value Then
                    Set myData = New Collection
                    myData.Add .Cells(i, "B").Value, "日付"
                    myData.Add WorksheetFunction.CountIf(.Columns("B"), .Cells(i, "B").Value), "件数"
                    myDataTable.Add myData
                    Set myData = Nothing
                End If
            End If
            
            '予定一覧のチェック欄にハイパーリンクを設定
            If .Cells(i, "F").Value = .Range("F2").Value Then
            
                .Cells(i, "F").Hyperlinks.Add Anchor:=.Cells(i, "F"), Address:="", _
                    SubAddress:=.Cells(i, "F").Address, TextToDisplay:="=" & .Range("F2").Address, ScreenTip:="クリックしてください"
            
            Else
                
                .Cells(i, "F").Hyperlinks.Add Anchor:=.Cells(i, "F"), Address:="", _
                    SubAddress:=.Cells(i, "F").Address, TextToDisplay:="　　", ScreenTip:="クリックしてください"
                
            End If
                
            
            i = i + 1
            
        Loop Until .Cells(i, "B").Value = ""
       
        
        
        '------- カレンダーシート作成 -------
        
        Set CalendarSheet = Create_CalendarSheet(CDate(myDataTable(1)(1)), CDate(myDataTable(myDataTable.Count)(1)))
        
        '行数調整
        For i = 1 To myDataTable.Count
        
            TargetDate = myDataTable(i)(1)
        
            Set FindRange = CalendarSheet.Cells.Find(What:=Format(TargetDate, "m月d日 aaa曜日"), LookIn:=xlValues, Lookat:=xlWhole)
            
            RowCount = FindRange.End(xlDown).Row - FindRange.Row - 1
            
            If RowCount < myDataTable(i)(2) Then
                
                For j = RowCount + 1 To myDataTable(i)(2)
                
                    FindRange.Offset(2).EntireRow.Copy
                    FindRange.Offset(2).EntireRow.Insert
                    Application.CutCopyMode = False
                    
                Next j
                
            End If

        Next i

        Set myDataTable = Nothing
        
        
        '------- 予定を日付ごとにコレクション化 -------
        i = 3
        
        Set myTableCollection = New Collection
        Set myDataTable = New Collection
    
        Do
        
            If IsDate(.Cells(i, "B").Value) Then
            
                Set myData = New Collection
                 
                myData.Add .Cells(i, "B").Value, "日付"
                myData.Add .Name, "シート名"
                myData.Add .Cells(i, "D").Address, "タスクアドレス"
                myData.Add .Cells(i, "C").Address, "時刻アドレス"
                myData.Add .Cells(i, "E").Address, "備考アドレス"
                myData.Add .Cells(i, "F").Address, "チェックアドレス"
                myDataTable.Add myData
                 
                Set myData = Nothing
                
                If .Cells(i, "B").Value <> .Cells(i + 1, "B").Value Then
                    
                    myTableCollection.Add myDataTable
                    Set myDataTable = Nothing
                    If .Cells(i + 1, "B").Value <> "" Then Set myDataTable = New Collection
                    
                End If
                
            End If
        
            i = i + 1
            
        Loop Until .Cells(i, "B").Value = ""
            
            
        '------- 予定をカレンダーに書き込み -------
        For i = 1 To myTableCollection.Count
        
            TargetDate = myTableCollection(i)(1)("日付")
            
            Set FindRange = CalendarSheet.Cells.Find(What:=Format(TargetDate, "m月d日 aaa曜日"), LookIn:=xlValues, Lookat:=xlWhole)
            
            For j = 1 To myTableCollection(i).Count
                
                Set TargetData = myTableCollection(i)(j)
                
                For k = 3 To TargetData.Count
                    
                    If k <> TargetData.Count Then
                        
                        TargetAddress = TargetData("シート名") & "!" & TargetData(k)
                        
                        FindRange.Offset(1).Cells(j, k - 2).Value = "=HYPERLINK(""#" & TargetAddress & """,if(" & TargetAddress & "="""",""""," & TargetAddress & "))"
                        FindRange.Offset(1).Cells(j, k - 2).Font.Color = rgbDarkBlue  '転記した予定の文字色
                        FindRange.Offset(1).Cells(j, k - 2).Font.Bold = True
                    
                    Else
                     
                        TargetAddress = TargetData("シート名") & "!" & TargetData(k)
                        
                        FindRange.Offset(1).Cells(j, k - 2).Hyperlinks.Add _
                           Anchor:=FindRange.Offset(1).Cells(j, k - 2), _
                           Address:="", _
                           SubAddress:=TargetAddress, _
                           TextToDisplay:="=if(" & TargetAddress & "="""",""　""," & TargetAddress & ")", _
                           ScreenTip:="クリックしてください"
                
                    End If
                
                Next k
                
            Next j

        Next i
        
    End With
    
    
    '------- ハイパーリンクの書式変更 -------
    For i = 1 To ThisWorkbook.Styles.Count
        
        If ThisWorkbook.Styles(i).Name Like "*Hyperlink" Then
        
            With ThisWorkbook.Styles(i).Font
                
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
            
            End With
            
        End If

    Next i

    CalendarSheet.Protect
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
