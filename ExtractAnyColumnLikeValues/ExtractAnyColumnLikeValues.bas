Attribute VB_Name = "Module1"
Option Explicit

Public Sub ExtractAnyColumnLikeValues()

Dim FindWords() As String
Dim TargetWords(1) As String
Dim TableRange As Range
Dim i As Long
Dim j As Long
Dim k As Long
Dim iRow As Long
Dim ChackFlag As Boolean
Dim ChackCount As Long
Dim FindColumn As Long
Dim TargetRange As Range
Dim FindRange As Range
Dim FindStart As Long

    Application.ScreenUpdating = False

    With ThisWorkbook.Sheets("食事記録")
        
        FindWords = Split(.Range("B1").Value, " ", , vbTextCompare)
        
        .Rows.Hidden = False
        
        With .Range("A4") '表の開始セル
            iRow = .Row + 1
            Set TableRange = .CurrentRegion
            With TableRange
                .Offset(1).Interior.ColorIndex = xlNone
                .Offset(1).Font.ColorIndex = xlAutomatic
                .Offset(1).Font.Bold = False
             End With
        End With

        Do
            
            ChackCount = 0
            
            Set TargetRange = Intersect(.Rows(iRow), TableRange)

            For i = 0 To UBound(FindWords)

                TargetWords(0) = StrConv(FindWords(i), vbWide)
                TargetWords(1) = StrConv(FindWords(i), vbNarrow)
                ChackFlag = False
                
                For j = 0 To 1
                    
                    If WorksheetFunction.CountIf(TargetRange, "*" & TargetWords(j) & "*") > 0 Then
                        
                        If ChackFlag = False Then
                            ChackCount = ChackCount + 1
                            ChackFlag = True
                        End If
                        
                        FindColumn = 0
                        
                        For k = 1 To WorksheetFunction.CountIf(TargetRange, "*" & TargetWords(j) & "*")
                            
                            FindColumn = WorksheetFunction.Match("*" & TargetWords(j) & "*", TargetRange.Offset(0, FindColumn), 0) + FindColumn
                            Set FindRange = .Cells(iRow, FindColumn)
                            FindRange.Interior.Color = vbYellow
                            
                            FindStart = 1
                            Do
                                With FindRange.Characters(Start:=InStr(FindStart, FindRange.Value, TargetWords(j), vbTextCompare), Length:=Len(TargetWords(j)))
                                    .Font.Color = vbRed
                                    .Font.Bold = True
                                End With
                                FindStart = InStr(FindStart, FindRange.Value, TargetWords(j), vbTextCompare) + Len(TargetWords(j))
                            Loop Until InStr(FindStart, FindRange.Value, TargetWords(j), vbTextCompare) = 0
    
                        Next k
       
                    End If
                
                Next j
                
            Next i
            
            Select Case .Range("B2").Value
                
                Case "OR"
                    
                    If ChackCount = 0 Then .Cells(iRow, "F").EntireRow.Hidden = True
            
                Case "AND"
                    
                    If ChackCount <> UBound(FindWords) + 1 Then .Cells(iRow, "F").EntireRow.Hidden = True
                   
            End Select
   
            iRow = iRow + 1
        
        Loop Until .Cells(iRow, "A").Value = ""
        
    End With
    
    Application.ScreenUpdating = True

End Sub



Public Sub ExtarctClear()
     Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("食事記録")
    
        .Rows.Hidden = False
        
        With .Range("A4")
            .CurrentRegion.Offset(1).Interior.ColorIndex = xlNone
            .CurrentRegion.Offset(1).Font.ColorIndex = xlAutomatic
            .CurrentRegion.Offset(1).Font.Bold = False
        End With
        
        
    End With
     Application.ScreenUpdating = True
     
End Sub
