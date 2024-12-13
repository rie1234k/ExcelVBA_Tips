Attribute VB_Name = "A_�\��]�L"
Option Explicit


Public Sub Write_Calendar()

Dim i As Long
Dim j As Long
Dim k As Long

Dim endRow As Long
Dim TargetRange As Range
Dim RowCount As Long
Dim TargetDate As Date
Dim CalendarSheet As Worksheet
Dim myData As Collection
Dim myDataTable As Collection
Dim myTableCollection As Collection
Dim FindRange As Range
Dim TargetData As Collection
Dim TargetAddress As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    With ThisWorkbook.Sheets("�\��ꗗ")
    
        '------- �\��ꗗ ���ёւ� -------
    
        endRow = .Cells(Rows.Count, "B").Row
        Set TargetRange = .Range(.Range("B2"), .Cells(endRow, "F"))
        
        If WorksheetFunction.CountA(.Columns("B")) > 3 Then
                .Sort.SortFields.Clear
                .Sort.SortFields.Add Key:=.Range("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .Sort.SortFields.Add Key:=.Range("C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .Sort.SetRange TargetRange
            With .Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
        
        '------- �J�����_�[�s�������p ���t�����擾 -------
        i = 3
        
        Set myDataTable = New Collection
        
        Do
            If IsDate(.Cells(i, "B").Value) Then
                
                If .Cells(i, "B").Value <> .Cells(i - 1, "B").Value Then
                    Set myData = New Collection
                    myData.Add .Cells(i, "B").Value, "���t"
                    myData.Add WorksheetFunction.CountIf(.Columns("B"), .Cells(i, "B").Value), "����"
                    myDataTable.Add myData
                    Set myData = Nothing
                End If
            End If
            
            '�\��ꗗ�̃`�F�b�N���Ƀn�C�p�[�����N��ݒ�
            If .Cells(i, "F").Value = .Range("F2").Value Then
            
                .Cells(i, "F").Hyperlinks.Add Anchor:=.Cells(i, "F"), Address:="", _
                    SubAddress:=.Cells(i, "F").Address, TextToDisplay:="=" & .Range("F2").Address, ScreenTip:="�N���b�N���Ă�������"
            
            Else
                
                .Cells(i, "F").Hyperlinks.Add Anchor:=.Cells(i, "F"), Address:="", _
                    SubAddress:=.Cells(i, "F").Address, TextToDisplay:="�@�@", ScreenTip:="�N���b�N���Ă�������"
                
            End If
                
            
            i = i + 1
            
        Loop Until .Cells(i, "B").Value = ""
       
        
        '------- �J�����_�[�V�[�g�쐬 -------
        
        Set CalendarSheet = Create_CalendarSheet(CDate(myDataTable(1)(1)), CDate(myDataTable(myDataTable.Count)(1)))
        
        '�s������
        For i = 1 To myDataTable.Count
        
            TargetDate = myDataTable(i)(1)
        
            Set FindRange = CalendarSheet.Cells.Find(What:=Format(TargetDate, "m��d�� aaa�j��"), LookIn:=xlValues, Lookat:=xlWhole)
            
            RowCount = FindRange.End(xlDown).Row - FindRange.Row - 1
            
            If RowCount < myDataTable(i)(2) Then
                
                For j = RowCount + 1 To myDataTable(i)(2)
                
                    FindRange.Offset(2).EntireRow.Copy
                    FindRange.Offset(2).EntireRow.Insert
                    Application.CutCopyMode = False
                    
                Next j
                
            End If

        Next i

        Set myDataTable = Nothing
        
        
        '------- �\�����t���ƂɃR���N�V������ -------
        i = 3
        
        Set myTableCollection = New Collection
        Set myDataTable = New Collection
    
        Do
        
            If IsDate(.Cells(i, "B").Value) Then
            
                Set myData = New Collection
                 
                myData.Add .Cells(i, "B").Value, "���t"
                myData.Add .Name, "�V�[�g��"
                myData.Add .Cells(i, "D").Address, "�^�X�N�A�h���X"
                myData.Add .Cells(i, "C").Address, "�����A�h���X"
                myData.Add .Cells(i, "E").Address, "���l�A�h���X"
                myData.Add .Cells(i, "F").Address, "�`�F�b�N�A�h���X"
                myDataTable.Add myData
                 
                Set myData = Nothing
                
                If .Cells(i, "B").Value <> .Cells(i + 1, "B").Value Then
                    
                    myTableCollection.Add myDataTable
                    Set myDataTable = Nothing
                    If .Cells(i + 1, "B").Value <> "" Then Set myDataTable = New Collection
                    
                End If
                
            End If
        
            i = i + 1
            
        Loop Until .Cells(i, "B").Value = ""
            
            
        '------- �\����J�����_�[�ɏ������� -------
        For i = 1 To myTableCollection.Count
        
            TargetDate = myTableCollection(i)(1)("���t")
            
            Set FindRange = CalendarSheet.Cells.Find(What:=Format(TargetDate, "m��d�� aaa�j��"), LookIn:=xlValues, Lookat:=xlWhole)
            
            For j = 1 To myTableCollection(i).Count
                
                Set TargetData = myTableCollection(i)(j)
                
                For k = 3 To TargetData.Count
                    
                    If k <> TargetData.Count Then
                        
                        TargetAddress = TargetData("�V�[�g��") & "!" & TargetData(k)
                        
                        FindRange.Offset(1).Cells(j, k - 2).Value = "=HYPERLINK(""#" & TargetAddress & """,if(" & TargetAddress & "="""",""""," & TargetAddress & "))"
                        FindRange.Offset(1).Cells(j, k - 2).Font.Color = rgbDarkBlue  '�]�L�����\��̕����F
                        FindRange.Offset(1).Cells(j, k - 2).Font.Bold = True
                    
                    Else
                     
                        TargetAddress = TargetData("�V�[�g��") & "!" & TargetData(k)
                        
                        FindRange.Offset(1).Cells(j, k - 2).Hyperlinks.Add _
                           Anchor:=FindRange.Offset(1).Cells(j, k - 2), _
                           Address:="", _
                           SubAddress:=TargetAddress, _
                           TextToDisplay:="=if(" & TargetAddress & "="""",""�@""," & TargetAddress & ")", _
                           ScreenTip:="�N���b�N���Ă�������"
                
                    End If
                
                Next k
                
            Next j

        Next i
        
    End With
    
    
    '------- �n�C�p�[�����N�̏����ύX -------
    With ThisWorkbook.Styles("Hyperlink").Font
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    With ThisWorkbook.Styles("Followed Hyperlink").Font
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With

    CalendarSheet.Protect
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
