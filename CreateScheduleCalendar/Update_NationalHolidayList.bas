Attribute VB_Name = "C_�j�����X�g�쐬"
Option Explicit

'���t�{ �u�����̏j���v�ɂ��� �����̏j��CSV�t�@�C��
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

    
    HolidaySheetName = "�j�����X�g"
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        
        If ThisWorkbook.Worksheets(i).Name = HolidaySheetName Then
            
            Set TargetSheet = ThisWorkbook.Worksheets(i)
            TargetYear = Year(TargetSheet.Range("A1").End(xlDown).Value) + 1
        
        End If
        
    Next i
    
    '�V�[�g���Ȃ��ꍇ�A�V�K�쐬
    If TargetSheet Is Nothing Then
        
        Set TargetSheet = Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        
        TargetSheet.Name = HolidaySheetName
    
        TargetSheet.Range("A1").Value = "���t"
        TargetSheet.Range("B1").Value = "����"
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
        ThisWorkbook.Names.Add Name:="�j�����X�g", RefersTo:="=" & Replace(TargetRange.Address(external:=True), "[" & ThisWorkbook.Name & "]", "")
        
    
    End With
    
    ThisWorkbook.Sheets(1).Activate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
  
End Sub

