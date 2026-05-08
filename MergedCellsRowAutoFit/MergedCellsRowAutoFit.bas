Attribute VB_Name = "Module1"
Option Explicit

Public Sub MergedCellsRowAutoFit()

Dim iShape As Shape
Dim TargetRange As Range
Dim iRange As Range
Dim AdjustHeight As Double

Const MAX_ROW_HEIGHT As Double = 409.5
Const SHAPE_MARGIN As Double = 5

    Application.ScreenUpdating = False
    On Error GoTo Termination

    With ActiveSheet
        Set iShape = .Shapes.AddLabel(msoTextOrientationHorizontal, 100, 100, 100, 100)
        Set TargetRange = .UsedRange
    End With
    
    With iShape.TextFrame2
        .MarginTop = SHAPE_MARGIN
        .MarginBottom = SHAPE_MARGIN
        .MarginLeft = SHAPE_MARGIN
        .MarginRight = SHAPE_MARGIN
    End With
    
    'êÊÇ…åãçáÇ≥ÇÍÇƒÇ¢Ç»Ç¢ÉZÉãÇÃçsÇÃçÇÇ≥Çé©ìÆí≤êÆÇ∑ÇÈ
    TargetRange.EntireRow.AutoFit
    
    For Each iRange In TargetRange
        
        If iRange.MergeCells And iRange.Address = iRange.MergeArea.Cells(1, 1).Address And iRange.WrapText = True Then
        
            If iRange.Value <> "" Then
                
                iShape.TextFrame2.TextRange.Text = iRange.Value
                
                iShape.Width = iRange.MergeArea.Width
                
                With iShape.TextFrame2.TextRange.Font
                    .Name = iRange.Font.Name
                    .NameFarEast = iRange.Font.Name
                    .Size = iRange.Font.Size
                End With
                
                iShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                
                If iRange.MergeArea.Height < iShape.Height Then
            
                    AdjustHeight = iRange.RowHeight + iShape.Height - iRange.MergeArea.Height
                
                    If AdjustHeight <= MAX_ROW_HEIGHT Then
                        iRange.RowHeight = AdjustHeight
                    Else
                        iRange.RowHeight = MAX_ROW_HEIGHT
                    End If
                    
                End If
            
            End If
        
        End If
        
        iShape.TextFrame2.TextRange.Text = ""
    
    Next iRange
    
Termination:
    On Error Resume Next
    If Not iShape Is Nothing Then iShape.Delete
    Set iShape = Nothing
    Application.ScreenUpdating = True
   
End Sub
