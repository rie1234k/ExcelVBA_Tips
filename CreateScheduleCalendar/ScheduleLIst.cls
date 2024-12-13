VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

Dim TargetNo As Long
Dim TargetSheetName As String
Dim TargetRange As Range
Dim mySheet As Worksheet
Dim myDic As Object

    If Cells(2, Target.Column).Value = "日付" And Target.Row > 2 And IsDate(Target.Value) Then
        
        TargetNo = Target.Row - WorksheetFunction.Match(CLng(Target.Value), Columns("B"), 0) + 1
        TargetSheetName = "カレンダー"

        Set myDic = CreateObject("Scripting.Dictionary")
        
        For Each mySheet In ThisWorkbook.Worksheets
            
           myDic.Add mySheet.Name, mySheet.Name
        
        Next mySheet
            
        If myDic.exists(TargetSheetName) Then
        
            Sheets(TargetSheetName).Activate
            
            Set TargetRange = ActiveSheet.Cells.Find(What:=Format(Target.Value, "m月d日 aaa曜日"), LookIn:=xlValues, Lookat:=xlWhole)
            
            If TargetRange Is Nothing Then
                
                Cancel = True
                Exit Sub
            
            End If
             
            TargetRange.Offset(TargetNo).Select
            
        End If
        
        Cancel = True
         
    End If

End Sub
