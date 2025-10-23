Attribute VB_Name = "C_祝日リスト作成"
Option Explicit

'内閣府 「国民の祝日」について 国民の祝日CSVファイル
'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv


Public Sub Update_NationalHolidayList()

Dim i As Long
Dim HolidaySheetName As String
Dim TargetSheet As Worksheet
Dim FSO As Object
Dim TargetYear As Long
Dim TargetFilePath As String
Dim TargetLine As Variant
Dim endRow As Long
Dim TargetRange As Range
Dim StartChar As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    
    HolidaySheetName = "祝日リスト"
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        
        If ThisWorkbook.Worksheets(i).Name = HolidaySheetName Then
            
            Set TargetSheet = ThisWorkbook.Worksheets(i)
            TargetYear = Year(WorksheetFunction.Max(TargetSheet.Columns(1))) + 1
        
        End If
        
    Next i
    
    'シートがない場合、新規作成
    If TargetSheet Is Nothing Then
        
        Set TargetSheet = Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        
        TargetSheet.Name = HolidaySheetName
    
        TargetSheet.Range("A1").Value = "日付"
        TargetSheet.Range("B1").Value = "名称"
        TargetYear = Year(Date) - 2
        
    End If
    
    TargetFilePath = GetDownloadFilePath("https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv")
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
     
    With FSO.GetFile(TargetFilePath).OpenAsTextStream
     
        Do
            TargetLine = Split(.ReadLine, ",")
            
            If Val(Left(TargetLine(0), 4)) >= TargetYear Then
            
                endRow = TargetSheet.Cells(Rows.Count, "A").End(xlUp).Offset(1).Row
                TargetSheet.Cells(endRow, "A").Value = TargetLine(0)
                TargetSheet.Cells(endRow, "B").Value = TargetLine(1)
                   
            End If
        
        Loop Until .AtEndOfStream
        
        .Close
    
    End With
    
    Set FSO = Nothing
    
    Kill TargetFilePath
    
    With TargetSheet
    
        .Columns("A:A").NumberFormatLocal = "yyyy/mm/dd (aaa) "
        .Columns("A:B").EntireColumn.AutoFit
        
        Set TargetRange = .Range(.Range("A1"), .Range("A1").End(xlDown))
        ThisWorkbook.Names.Add Name:="祝日リスト", RefersTo:="=" & Replace(TargetRange.Address(external:=True), "[" & ThisWorkbook.Name & "]", "")
        
    
    End With
    
    ThisWorkbook.Sheets(1).Activate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
  
End Sub

