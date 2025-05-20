Attribute VB_Name = "Module1"
Option Explicit

Public Sub ExtractAnyColumnLikeValues()
Dim FindWords As Variant
Dim i As Long
Dim j As Long
Dim iRow As Long
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
            .CurrentRegion.Offset(1).Interior.ColorIndex = xlNone
            .CurrentRegion.Offset(1).Font.ColorIndex = xlAutomatic
            .CurrentRegion.Offset(1).Font.Bold = False
             iRow = .Row + 1
        End With

        Do
            
            ChackCount = 0
            
            Set TargetRange = Intersect(.Rows(iRow), .Range("A4").CurrentRegion)

            For i = 0 To UBound(FindWords)
                
                FindColumn = 0
                
                If WorksheetFunction.CountIf(TargetRange, "*" & FindWords(i) & "*") > 0 Then
                    
                    ChackCount = ChackCount + 1
                    
                    For j = 1 To WorksheetFunction.CountIf(TargetRange, "*" & FindWords(i) & "*")
                        
                        FindColumn = WorksheetFunction.Match("*" & FindWords(i) & "*", TargetRange.Offset(0, FindColumn), 0) + FindColumn
                        Set FindRange = .Cells(iRow, FindColumn)
                        FindRange.Interior.Color = vbYellow
                        
                        FindStart = 1
                        Do
                            With FindRange.Characters(Start:=InStr(FindStart, FindRange.Value, FindWords(i)), Length:=Len(FindWords(i)))
                                .Font.Color = vbRed
                                .Font.Bold = True
                            End With
                            FindStart = InStr(FindStart, FindRange.Value, FindWords(i)) + Len(FindWords(i))
                        Loop Until InStr(FindStart, FindRange.Value, FindWords(i)) = 0

                    Next j
   
                End If
                
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
