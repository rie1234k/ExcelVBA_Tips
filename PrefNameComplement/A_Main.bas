Attribute VB_Name = "A_Main"
Option Explicit

Public Sub PrefNameComplement()

Dim TargetUrl As String
Dim PostcodeFileName As String

Dim i As Long
Dim j As Long
Dim k As Long

Dim ListCollection As Collection
Dim CityList() As String
Dim TownList() As String
Dim DupCityDic As Object
Dim DupTownDic As Object

Dim TargetSheet As Worksheet
Dim endRow As Long
Dim OutputRange As Range

Dim AddressList() As String
Dim SearchList()  As String
Dim SearchListCount As Long
Dim SearchListWord(2) As String
Dim SearchListFlag(2) As Boolean
Dim SearchTownName As String


    
    
    '�X�֔ԍ��f�[�^�_�E�����[�h���Z���̗X�֔ԍ��iCSV�`���j���ǂ݉����f�[�^�̑����E�X�����������ŕ\�L������́��S���ꊇ
    TargetUrl = "https://www.post.japanpost.jp/zipcode/dl/kogaki/zip/ken_all.zip"
    
    '�_�E�����[�h����Zip�t�@�C������CSV�t�@�C����
    PostcodeFileName = "KEN_ALL.CSV"
    
     Application.StatusBar = "�f�[�^���_�E�����[�h���ł�..."
    
    Call PostcodeZipFileDowunload(TargetUrl, PostcodeFileName)

    
    Application.StatusBar = "�s�撬�����X�g�E���惊�X�g���쐬���ł�..."
    
    Set ListCollection = New Collection
    Set ListCollection = GetListCollection(PostcodeFileName)
    
    CityList = ListCollection("CityList")
    TownList = ListCollection("TownList")
    Set DupTownDic = ListCollection("DupTownDic")
    Set DupCityDic = ListCollection("DupCityDic")
    
    Set ListCollection = Nothing

    Application.StatusBar = "�����p���[�h���X�g���쐬���ł�..."
    
    Set TargetSheet = ThisWorkbook.Sheets("��ƃV�[�g")
    endRow = TargetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    
     '------- �ΏۏZ�����X�g��z��Ƃ��Ď擾 -------
    ReDim AddressList(2 To endRow)
    
    For i = 2 To endRow
        
        AddressList(i) = TargetSheet.Cells(i, "A").Value

    Next i
    

    '------- �����p���[�h���X�g�쐬 -------
    For i = 0 To UBound(CityList, 2)
    
        SearchListWord(0) = CityList(3, i)
        SearchListWord(1) = CityList(4, i)
        SearchListWord(2) = CityList(5, i)
        
        '�s�������ƌS+�s�������������ꍇ�A�s�������݂̂̌����ł悢�̂ŁA�󗓂Ƃ���
        If CityList(3, i) = CityList(5, i) Then SearchListWord(2) = ""
        
        For j = 0 To 2
        
            SearchListFlag(j) = False
        
        Next j
        
        
        For j = 0 To 2
        
            k = 2
            
            Do
                If AddressList(k) Like SearchListWord(j) & "*" _
                    And SearchListWord(j) <> "" Then

                    SearchListFlag(j) = True
                    
                End If
                
                If k = UBound(AddressList) Then Exit Do
                
                k = k + 1
            
            Loop Until SearchListFlag(j)
        
        Next j
        
        
        For j = 0 To 2
        
            If SearchListFlag(j) Then
            
                ReDim Preserve SearchList(6, SearchListCount)
                
                For k = 0 To UBound(SearchList, 1) - 1
                
                    SearchList(k, SearchListCount) = CityList(k, i)
                
                Next k
                
                SearchList(UBound(SearchList, 1), SearchListCount) = SearchListWord(j)
                SearchListCount = SearchListCount + 1
            
            End If
        
        Next j
    
    Next i
    
    '------- �s���{���E�s�撬�����擾���� -------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
     '�h��Ԃ�����
    TargetSheet.Range("A1").CurrentRegion.Interior.Color = xlNone

    '�I�[�g�t�B���^�[���ݒ肳��Ă���ꍇ�ɂ͉���
    If Not TargetSheet.AutoFilter Is Nothing Then TargetSheet.Range("A1").AutoFilter
    
    TargetSheet.Range("A1").CurrentRegion.Offset(1, 1).ClearContents
    
    
    For i = 0 To UBound(SearchList, 2)
        
        With TargetSheet
            
            Select Case True
                
                '�������[�h�Ǝs���������������ɏd���s�撬�����̏ꍇ�͒���Ō���
                Case SearchList(6, i) = SearchList(3, i) And DupCityDic.Exists(SearchList(4, i))
                    
                    For j = 0 To UBound(TownList, 2)
                        
                        If TownList(3, j) = SearchList(3, i) Then
                        
                            SearchTownName = TownList(3, j) & "*" & TownList(4, j)
                            
                            Select Case True
                                
                                '�����̒��悪����ꍇ
                                Case DupTownDic.Exists(SearchTownName)
                            
                                   If WorksheetFunction.CountIf(.Columns("A:A"), SearchTownName & "*") > 0 Then
    
                                    .Range("A1").AutoFilter Field:=1, Criteria1:=SearchTownName & "*"
                                    Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
                                    OutputRange.SpecialCells(xlCellTypeVisible).Value = "�������̒��悪���邽�ߗv�m�F��"
                                    .Range("A1").AutoFilter
                                    
                                    End If
                                
                                Case Else
                                    
                                    If WorksheetFunction.CountIf(.Columns("A:A"), SearchTownName & "*") > 0 Then
                                
                                    .Range("A1").AutoFilter Field:=1, Criteria1:=SearchTownName & "*"
                                    
                                    For k = 0 To 3
                                        
                                        Set OutputRange = .Range(.Cells(2, k + 2), .Cells(endRow, k + 2))
                                        OutputRange.SpecialCells(xlCellTypeVisible).Value = TownList(k, j)
                                
                                    Next k
                                
                                    .Range("A1").AutoFilter
                                
                                End If
                            
                            End Select
                            
                        End If
                        
                    Next j
                    
                Case Else
                    
                    .Range("A1").AutoFilter Field:=1, Criteria1:=SearchList(6, i) & "*"
                
                    For j = 0 To 3
                        
                        Set OutputRange = .Range(.Cells(2, j + 2), .Cells(endRow, j + 2))
                        OutputRange.SpecialCells(xlCellTypeVisible).Value = SearchList(j, i)
                    
                    Next j
                    
                    .Range("A1").AutoFilter


            End Select
            
            
        End With
        
        Application.StatusBar = i & "/" & UBound(SearchList, 2) & "���ڂ��������ł�..."
        
    
    Next i
    

    Application.StatusBar = "���X�g���������ł�..."
    
    '------- �v�m�F�Z���Ƀ}�[�J�[ -------
    
    With TargetSheet

        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:=""
            Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
            OutputRange.SpecialCells(xlCellTypeVisible).Value = "���Y������f�[�^���Ȃ����ߗv�m�F��"
        
        End If
        
        .Range("A1").AutoFilter
        
        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "��*") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:="��*"
            Set OutputRange = .Range(.Range("A2"), .Cells(endRow, "A"))
            OutputRange.SpecialCells(xlCellTypeVisible).Interior.Color = vbYellow
        
        End If
        
        .Range("A1").AutoFilter
        
        
        '�Z���⊮
        Call ComplementAddress(TargetSheet)
    
    
        .Range("A1").AutoFilter
    
    End With
    
    Set TargetSheet = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
 
    
End Sub


Public Sub ComplementAddress(TargetSheet As Worksheet)

Dim endRow As Long
Dim ReplaceAddress As String
Dim FullAddress As String
Dim i As Long
Dim j As Long

    endRow = TargetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    With TargetSheet
    
        For i = 2 To endRow
        
            If .Cells(i, "C").Value <> "" Then
                
                ReplaceAddress = .Cells(i, "A").Value
                
                For j = 3 To 5
                 
                    ReplaceAddress = Replace(ReplaceAddress, .Cells(i, j).Value, "")
                 
                Next j
                
                .Cells(i, "F").Value = ReplaceAddress
                
                FullAddress = ""
                
                For j = 3 To 6
                    
                    FullAddress = FullAddress & .Cells(i, j).Value
                
                Next j

                .Cells(i, "G").Value = FullAddress
                
            End If
        
        Next i
        
    End With

End Sub
