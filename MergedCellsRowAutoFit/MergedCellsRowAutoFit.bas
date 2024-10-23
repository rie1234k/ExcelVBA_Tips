Attribute VB_Name = "Module1"
Option Explicit

Public Sub MergedCellsRowAutoFit()

Dim obj As Shape
Dim TargetRange As Range
Dim iRange As Range
Dim AdjustHeight As Double

    Application.ScreenUpdating = False

    With ActiveSheet
        
        Set obj = .Shapes.AddLabel(msoTextOrientationHorizontal, 100, 100, 100, 100)
        Set TargetRange = .Range(.Range("A1"), .Range("A1").SpecialCells(xlCellTypeLastCell))
    
    End With
     
    obj.TextFrame2.MarginTop = 5
    obj.TextFrame2.MarginBottom = 5
    obj.TextFrame2.MarginLeft = 5
    obj.TextFrame2.MarginRight = 5

    TargetRange.EntireRow.AutoFit
    
    For Each iRange In TargetRange
    
        If iRange.MergeArea.Count = 1 Then
        
            If iRange.Value <> "" Then
            
                obj.TextFrame2.TextRange.Text = iRange.Value
            
            End If
        
        Else
        
            If iRange.MergeArea.Value2(1, 1) <> "" Then
            
                obj.TextFrame2.TextRange.Text = iRange.MergeArea.Value2(1, 1)
            
            End If
        
        End If
        
        If obj.TextFrame2.TextRange.Text <> "" Then
        
            obj.Width = iRange.MergeArea.Width
            
            obj.TextFrame2.TextRange.Font.Name = iRange.Font.Name
            obj.TextFrame2.TextRange.Font.NameFarEast = iRange.Font.Name
            obj.TextFrame2.TextRange.Font.Size = iRange.Font.Size
            obj.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            
            If iRange.MergeArea.Height < obj.Height Then
            
                AdjustHeight = iRange.RowHeight + obj.Height - iRange.MergeArea.Height
                
                If AdjustHeight <= 409.5 Then

                    iRange.RowHeight = AdjustHeight
                
                Else
                
                    iRange.RowHeight = 409.5
                
                End If
            
            End If
        
        End If
        
        obj.TextFrame2.TextRange.Text = ""
    
    Next iRange
    
    obj.Delete
    Set obj = Nothing
          
    Application.ScreenUpdating = True
    
End Sub
