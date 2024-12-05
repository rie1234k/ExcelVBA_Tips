Attribute VB_Name = "Module1"
Option Explicit


Public Sub Set_DynamicChartRange()

Dim TargetWorksheet As Worksheet
Dim StartItemRange  As Range
Dim EndItemRange As Range

Dim TargetSeriesCollection As Collection
Dim TargetChartObject As ChartObject
Dim TargetSeries As Series
Dim myItem As Collection
Dim TargetDirection As String

Dim StartFormula As String
Dim EndFormula As String
Dim CountFormula As String
Dim TargetAxesStr As String
Dim TargetAxesStartStr As String
Dim TargetSeriesStr As String
Dim TargetSeriesStartStr As String
Dim TargetSeriesIndex As String
Dim CharStart As Long
Dim CharLength As Long
Dim ChangeFormula As String
  
    '------- �ݒ�J�n -------
    Set TargetWorksheet = ActiveSheet
    Set StartItemRange = TargetWorksheet.Range("C3") '�J�n���ړ��̓Z����ݒ�
    Set EndItemRange = TargetWorksheet.Range("C4") '�I�����ړ��̓Z����ݒ�
    '------- �ݒ�I�� -------

     TargetWorksheet.Names.Add Name:=TargetWorksheet.Name & "_�͈͊J�n", RefersTo:="='" & TargetWorksheet.Name & "'!" & StartItemRange.Address
     TargetWorksheet.Names.Add Name:=TargetWorksheet.Name & "_�͈͏I��", RefersTo:="='" & TargetWorksheet.Name & "'!" & EndItemRange.Address
     
     Set TargetSeriesCollection = New Collection
                       
     For Each TargetChartObject In TargetWorksheet.ChartObjects
         
         For Each TargetSeries In TargetChartObject.Chart.FullSeriesCollection
            
            CharStart = InStr(TargetSeries.Formula, ",") + 1
            CharLength = InStr(CharStart, TargetSeries.Formula, ",") - CharStart
            TargetAxesStr = Mid(TargetSeries.Formula, CharStart, CharLength)
            
            CharStart = InStr(CharStart, TargetSeries.Formula, ",") + 1
            CharLength = InStr(CharStart, TargetSeries.Formula, ",") - CharStart
            TargetSeriesStr = Mid(TargetSeries.Formula, CharStart, CharLength)
            TargetSeriesIndex = Replace(Mid(TargetSeries.Formula, InStrRev(TargetSeries.Formula, ",") + 1), ")", "")
             
            '�A������f�[�^�̂ݑΏ�
            If Len(TargetSeries.Formula) - Len(Replace(TargetSeries.Formula, ",", "")) = 3 And InStr(TargetSeriesStr, ":") > 0 Then

                '�J�n�E�I�����ڂ����x���͈͂ɂ��邩
                If WorksheetFunction.CountIf(TargetWorksheet.Range(TargetAxesStr), StartItemRange.Value) _
                    And WorksheetFunction.CountIf(TargetWorksheet.Range(TargetAxesStr), EndItemRange.Value) Then
                    
                    '���x���̕����m�F
                    If TargetWorksheet.Range(TargetAxesStr).Columns.Count > TargetWorksheet.Range(TargetAxesStr).Rows.Count Then
                        
                        TargetDirection = "��"
                        
                    Else
                    
                        TargetDirection = "�c"
                        
                    End If

                    Set myItem = New Collection
                    
                    myItem.Add TargetSeries, "�n��"
                    myItem.Add TargetAxesStr, "�����x���͈�"
                    myItem.Add TargetWorksheet.Range(TargetAxesStr).Address(ReferenceStyle:=xlR1C1), "�����x���͈�R1C1"
                    myItem.Add TargetDirection, "�����x������"
                    myItem.Add TargetSeriesStr, "�n��͈�"
                    myItem.Add TargetWorksheet.Range(TargetSeriesStr).Address(ReferenceStyle:=xlR1C1), "�n��͈�R1C1"
                    myItem.Add "�n��͈�" & TargetSeriesIndex, "�n��"
                    myItem.Add Replace(TargetChartObject.Name, " ", ""), "�O���t��"
                    
                    TargetSeriesCollection.Add myItem
                    
                    Set myItem = Nothing
                
                End If
                
            End If
            
         Next TargetSeries
     
     Next TargetChartObject
 

     For Each myItem In TargetSeriesCollection
            
        '���x���͈͂̋N�_�Z���A�h���X
        TargetAxesStartStr = Left(myItem("�����x���͈�"), InStr(myItem("�����x���͈�"), ":") - 1)
        
        '���x��(�s�E��)�S�̖̂��O��`
        Select Case myItem("�����x������")
        
            Case "��"
            
                TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_�����x���͈͑S��", RefersTo:="=" & TargetWorksheet.Range(TargetAxesStartStr).EntireRow.Address
           
            Case "�c"
                
                TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_�����x���͈͑S��", RefersTo:="=" & TargetWorksheet.Range(TargetAxesStartStr).EntireColumn.Address
        
        End Select
        
        '�J�n�ʒu�A�\�������̖��O��`
        StartFormula = "MATCH(" & TargetWorksheet.Name & "_�͈͊J�n," & myItem("�O���t��") & "_�����x���͈͑S��" & ",0)"
        EndFormula = "MATCH(" & TargetWorksheet.Name & "_�͈͏I��," & myItem("�O���t��") & "_�����x���͈͑S��" & ",0)"
        CountFormula = EndFormula & " - " & StartFormula & " +1"

        TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_�J�n�ʒu", RefersTo:="=" & StartFormula
        TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_�\������", RefersTo:="=" & CountFormula
        
         '�n��͈͂̋N�_�Z���A�h���X
        TargetSeriesStartStr = Left(myItem("�n��͈�"), InStr(myItem("�n��͈�"), ":") - 1)
        
        '�w�莲���x���͈́A�n��̖��O��`
        Select Case myItem("�����x������")
              
              Case "��"
                
                TargetAxesStartStr = Replace(TargetAxesStartStr, Mid(TargetAxesStartStr, InStr(TargetAxesStartStr, "$") + 1), "A") & Mid(TargetAxesStartStr, InStrRev(TargetAxesStartStr, "$"))
                TargetSeriesStartStr = Replace(TargetSeriesStartStr, Mid(TargetSeriesStartStr, InStr(TargetAxesStartStr, "$") + 1), "A") & Mid(TargetSeriesStartStr, InStrRev(TargetSeriesStartStr, "$"))
                TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_�w�莲���x���͈�", RefersTo:="=OFFSET(" & TargetAxesStartStr & ",0," & myItem("�O���t��") & "_�J�n�ʒu -1,1," & myItem("�O���t��") & "_�\������)"
                TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_" & myItem("�n��"), RefersTo:="=OFFSET(" & TargetSeriesStartStr & ",0," & myItem("�O���t��") & "_�J�n�ʒu -1,1," & myItem("�O���t��") & "_�\������)"

            Case "�c"
            
                TargetAxesStartStr = Replace(TargetAxesStartStr, Mid(TargetAxesStartStr, InStrRev(TargetAxesStartStr, "$") + 1), "1")
                TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_�w�莲���x���͈�", RefersTo:="=OFFSET(" & TargetAxesStartStr & "," & myItem("�O���t��") & "_�J�n�ʒu -1,0," & myItem("�O���t��") & "_�\������,1)"
                TargetSeriesStartStr = Replace(TargetSeriesStartStr, Mid(TargetSeriesStartStr, InStrRev(TargetSeriesStartStr, "$") + 1), "1")
                TargetWorksheet.Names.Add Name:=myItem("�O���t��") & "_" & myItem("�n��"), RefersTo:="=OFFSET(" & TargetSeriesStartStr & "," & myItem("�O���t��") & "_�J�n�ʒu -1,0," & myItem("�O���t��") & "_�\������,1)"

        End Select
        
        ChangeFormula = Replace(myItem("�n��").FormulaR1C1, myItem("�����x���͈�R1C1"), myItem("�O���t��") & "_�w�莲���x���͈�")
        ChangeFormula = Replace(ChangeFormula, myItem("�n��͈�R1C1"), myItem("�O���t��") & "_" & myItem("�n��"))
        
        myItem("�n��").FormulaR1C1 = ChangeFormula
        
    Next myItem
     
    MsgBox "�O���t�ݒ肪�������܂���"
    
     
End Sub






