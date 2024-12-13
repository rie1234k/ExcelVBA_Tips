VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
Dim TargetRange As Range
Dim StartTime As Date
        
        StartTime = Now
        
        If Target.ScreenTip = "クリックしてください" Then
        
            Do
                If ActiveCell.Address(External:=True) Like "*" & Target.SubAddress Then
                    
                    Set TargetRange = ActiveCell
                    
                    Exit Do
                
                End If
                
                If StartTime + TimeValue("0:00:01") < Now Then Exit Sub
                
            Loop
             
            If TargetRange.Value = ActiveSheet.Range("F2").Value Then
                
                TargetRange.Value = "　　"
                TargetRange.Hyperlinks.Add Anchor:=TargetRange, Address:="", _
                    SubAddress:=TargetRange.Address, ScreenTip:="クリックしてください"
                
            Else
            
                TargetRange.Value = "=" & ActiveSheet.Range("F2").Address
                TargetRange.Hyperlinks.Add Anchor:=TargetRange, Address:="", _
                    SubAddress:=TargetRange.Address, ScreenTip:="クリックしてください"
                
            End If
            
            If Sh.Name <> ActiveSheet.Name Then Application.Wait [Now()] + 600 / 86400000
            
            Sh.Select
         
        End If
 
End Sub
