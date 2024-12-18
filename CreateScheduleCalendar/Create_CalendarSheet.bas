Attribute VB_Name = "B_�J�����_�[�쐬"
Option Explicit


Public Function Create_CalendarSheet(StartDate As Date, EndDate As Date) As Worksheet

Dim CalendarSheetName As String
Dim ListSheet As Worksheet
Dim TargetSheet As Worksheet
Dim InsertRowCount As Long
    
    Set ListSheet = ThisWorkbook.Sheets("�\��ꗗ")
    CalendarSheetName = "�J�����_�["
    
    For Each TargetSheet In ThisWorkbook.Worksheets
    
        If TargetSheet.Name = CalendarSheetName Then
        
            Application.DisplayAlerts = False
            TargetSheet.Delete
            Application.DisplayAlerts = True
            
        End If

    Next TargetSheet

    Set TargetSheet = Sheets.Add(After:=ThisWorkbook.Sheets(1))
    
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    TargetSheet.Name = CalendarSheetName
    
    TargetSheet.Range("B1").Value = DateSerial(Year(StartDate), Month(StartDate), 1)
    TargetSheet.Range("F1").Value = DateSerial(Year(EndDate), Month(EndDate), 1)
    TargetSheet.Range("B1", "F1").Font.Color = vbWhite
    
    TargetSheet.Range("B2:F2").Merge
    
    With TargetSheet.Range("B2")
        .Value = Format(StartDate, "yyyy�Nm�� �` ") & Format(EndDate, "yyyy�Nm��")
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignLeft
    End With
 
    TargetSheet.Rows("1:1").RowHeight = 3
    TargetSheet.Rows("3:3").RowHeight = 3
    TargetSheet.Columns("A:A").ColumnWidth = 2

    With TargetSheet.Range("B3:E3")
        .Value = 1
        .Font.Color = vbWhite
    End With

    With Range("F3")
        .FormulaR1C1 = "=RC[-4]+1"
        .Font.Color = vbWhite
        .AutoFill Destination:=Range("F3:AC3"), Type:=xlFillDefault
    End With
    
    With TargetSheet.Range("B4")
        .FormulaR1C1 = "=R1C2+R[-1]C-WEEKDAY(R1C2,2)" 'WEEKDAY ���:2 ���j�n�܂�
        .NumberFormatLocal = "m""��""d""�� ""aaa""�j��"""
    End With
    
    TargetSheet.Range("C4").Value = "��" & ListSheet.Name & "!" & ListSheet.Range("C2").Address
    TargetSheet.Range("D4").Value = "��" & ListSheet.Name & "!" & ListSheet.Range("E2").Address
    TargetSheet.Range("E4").Value = "��" & ListSheet.Name & "!" & ListSheet.Range("F2").Address
    
    
    With TargetSheet.Range("B4:E7")
        .BorderAround xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
        .Borders(xlInsideVertical).Weight = xlHairline
        .ShrinkToFit = True
    End With
    
    '�`�F�b�N����ꂽ�ꍇ�̏����t������
    With TargetSheet.Range("B5:D7")
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=OFFSET($B5,0,(B$3-1)*4+3,1,1)=�\��ꗗ!$F$2"
        .FormatConditions(1).Font.Color = rgbSilver
    End With
    
    '����/�y���Z��������F�h��
    With TargetSheet.Range("B4:E4")
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=OFFSET($B4,0,(B$3-1)*4,1,1)=TODAY()"
        .FormatConditions(1).Interior.Color = rgbGold
        .FormatConditions.Add xlExpression, Formula1:="=WEEKDAY(OFFSET($B4,0,(B$3-1)*4,1,1),2)=6"
        .FormatConditions(2).Interior.Color = rgbAliceBlue
        .FormatConditions.Add xlExpression, Formula1:="=WEEKDAY(OFFSET($B4,0,(B$3-1)*4,1,1),2)=7"
        .FormatConditions(3).Interior.Color = rgbMistyRose
        .FormatConditions.Add xlExpression, Formula1:="=COUNTIF(�j�����X�g,OFFSET($B4,0,(B$3-1)*4,1,1))=1"
        .FormatConditions(4).Interior.Color = rgbMistyRose
    End With
    
    
    TargetSheet.Range("C5:E7").HorizontalAlignment = xlCenter
    TargetSheet.Range("C5:C7").NumberFormatLocal = "h:mm;@"
    
    
    With TargetSheet.Range("B4:E4")
        .Font.Size = 12
        .Interior.Color = rgbLemonChiffon
        .HorizontalAlignment = xlCenter
        .BorderAround xlContinuous
    End With
    
    With TargetSheet
        .Range("B4:E7").Copy Range("B8")
        .Range("B8").FormulaR1C1 = "=R[-4]C+7"
        .Range("C8").FormulaR1C1 = "=R[-4]C"
        .Range("D8").FormulaR1C1 = "=R[-4]C"
        
        'End�v���p�e�B�ōs�����m�F����Ƃ��ɕK�v�Ȃ��߁A�Ō�̌��̗�����1�T�ԕ��܂ō쐬
        '���ɍŏ���1�T�ԕ��쐬�ς݂̂��߁A-1�Œ���
        InsertRowCount = WorksheetFunction.RoundUp((DateAdd("m", 1, TargetSheet.Range("F1").Value) + 7 - TargetSheet.Range("B4").Value) / 7, 0) - 1
        
        .Range("A8:E11").AutoFill Destination:=.Range(.Range("A8"), Cells(8 + InsertRowCount * 4 - 1, "E")), Type:=xlFillDefault
        .Range(.Range("B4"), Cells(8 + InsertRowCount * 4 - 1, "E")).AutoFill Destination:=.Range(.Range("B4"), Cells(8 + InsertRowCount * 4 - 1, "AC")), Type:=xlFillDefault
        .Columns("B:AC").EntireColumn.AutoFit
    
        
    End With
    
    Set Create_CalendarSheet = TargetSheet
    
End Function
