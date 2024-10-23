Attribute VB_Name = "Module1"
Option Explicit

Public Sub MergedCellsRowAutoFit()

Dim i As Long
Dim obj As Shape
Dim TargetRange As Range
Dim iRange As Range
Dim iHeight As Double


    Application.ScreenUpdating = False

    With ActiveSheet
        
        Set obj = .Shapes.AddLabel(msoTextOrientationHorizontal, 100, 100, 100, 100)
        
        obj.TextFrame2.MarginTop = 5
        obj.TextFrame2.MarginBottom = 5
        obj.TextFrame2.MarginLeft = 5
        obj.TextFrame2.MarginRight = 5

        Set TargetRange = .Range(.Range("A1"), .Range("A1").SpecialCells(xlCellTypeLastCell))
        
        TargetRange.EntireRow.AutoFit
        
        For i = 1 To TargetRange.Count
            
            Set iRange = TargetRange.Item(i)
            
            If iRange.MergeArea.Count = 1 Then
                
                If iRange.Value <> "" Then obj.TextFrame2.TextRange.Text = iRange.Value
                    
            Else
                
                If iRange.MergeArea.Value2(1, 1) <> "" Then obj.TextFrame2.TextRange.Text = iRange.MergeArea.Value2(1, 1)
                        
            End If
                       
            If obj.TextFrame2.TextRange.Text <> "" Then
                  
                obj.Width = iRange.MergeArea.Width
                
                obj.TextFrame2.TextRange.Font.Name = iRange.Font.Name
                obj.TextFrame2.TextRange.Font.NameFarEast = iRange.Font.Name
                obj.TextFrame2.TextRange.Font.Size = iRange.Font.Size
                obj.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                
                iHeight = obj.Height
                
                If iRange.MergeArea.Height < iHeight Then
                    
                    If iRange.RowHeight + iHeight - iRange.MergeArea.Height <= 409.5 Then
            
                        iRange.RowHeight = iRange.RowHeight + iHeight - iRange.MergeArea.Height
                    
                    Else
                            
                        iRange.RowHeight = 409.5
                            
                    End If
                    
                End If
            
            End If
            
            obj.TextFrame2.TextRange.Text = ""
            
            Set iRange = Nothing
            
        Next i

        obj.Delete
        
        Set obj = Nothing
     
    End With
                 
    Application.ScreenUpdating = True
    
    
End Sub
