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

Private Sub CreateBlankLink()

Dim i As Long
Dim myHyperlink  As Hyperlink
    
    i = 4
    
    Do
                   
        Cells(i, "E").Hyperlinks.Delete
        Cells(i, "E").Borders.LineStyle = xlContinuous
        Set myHyperlink = ActiveSheet.Hyperlinks.Add(Anchor:=Cells(i, "E"), Address:="", ScreenTip:="メールアドレスをクリック")
        
        i = i + 1
    
    Loop Until Cells(i, "E").Value = ""


End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

Dim Title As String
Dim Body As String
Dim myHyperlink  As Hyperlink
Dim TargetRow As Long

    
    TargetRow = ActiveCell.Row
    
    '件名に改行がある場合は削除し、全角の空白を半角の空白にした後、半角の空白を「%20」にエンコード
    '半角の「&」があるとうまく作動しないため、全角の「＆」に置き換え
    Title = Replace(Replace(Replace(WorksheetFunction.Clean(Cells(TargetRow, "D").Value), "　", " "), " ", "%20"), "&", "＆")
    
    '担当者に"様"をつけて、会社名担当者名を本文の文頭に追加
    Body = Cells(TargetRow, "H").Value & vbCrLf & Replace(Cells(TargetRow, "I"), Chr(10), " 様" & Chr(10)) & " 様" & vbCrLf & vbCrLf & Range("B4") & vbCrLf
 
    
    'クリップボードに文字列を格納
     With CreateObject("Forms.TextBox.1")
       .MultiLine = True
       .Text = Body
       .SelStart = 0
       .SelLength = .TextLength
       .Copy
     End With
    
      Set myHyperlink = Hyperlinks.Add(Anchor:=Cells(TargetRow, "E"), _
                        Address:="mailto:" & Cells(TargetRow, "E") & _
                        "?cc=" & Cells(TargetRow, "F") & _
                        "&bcc=" & Cells(TargetRow, "G") & _
                        "&subject=" & Title, _
                        TextToDisplay:=Cells(TargetRow, "E").Value, ScreenTip:="メールアドレスをクリック")
    
    
      Application.EnableEvents = False
      Selection.Hyperlinks(1).Follow
      Application.EnableEvents = True
      
      Set myHyperlink = ActiveSheet.Hyperlinks.Add(Anchor:=Cells(TargetRow, "E"), Address:="", ScreenTip:="メールアドレスをクリック")
      
      
      
End Sub

