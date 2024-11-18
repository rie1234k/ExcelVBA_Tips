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
Dim CityListNo As Collection
Dim TownListNo As Collection
Dim DupCityDic As Object
Dim DupTownDic As Object
Dim TypoDic As Object

Dim TargetSheet As Worksheet
Dim endRow As Long
Dim OutputRange As Range

Dim AddressList() As String
Dim SearchList()  As String
Dim SearchListCount As Long
Dim SearchListNo As Collection
Dim SearchWord() As String
Dim SearchFlag() As Boolean
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
    Set CityListNo = ListCollection("CityListNo")
    Set TownListNo = ListCollection("TownListNo")
    Set DupTownDic = ListCollection("DupTownDic")
    Set DupCityDic = ListCollection("DupCityDic")
    Set TypoDic = ListCollection("TypoDic")
    
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
    
    Set SearchListNo = CityListNo
    SearchListNo.Add 7, "SearchWord"
   
    For i = 0 To UBound(CityList, 2)
        
        ReDim SearchWord(1)
           
        SearchWord(0) = CityList(CityListNo("PrefCityName"), i)
        SearchWord(1) = CityList(CityListNo("CityName"), i)
     
     
        If CityList(CityListNo("AreaName"), i) <> "" Then
            
            ReDim Preserve SearchWord(3)
            
            SearchWord(2) = CityList(CityListNo("PrefAreaCityName"), i)
            SearchWord(3) = CityList(CityListNo("AreaCityName"), i)
        
        End If
        
               
        ReDim SearchFlag(UBound(SearchWord))
        
        For j = 0 To UBound(SearchFlag)
            
            SearchFlag(j) = False

        Next j
        
        
        For j = 0 To UBound(SearchWord)
               
            k = 2
        
            Do
                If AddressList(k) Like SearchWord(j) & "*" Then

                    SearchFlag(j) = True
                    
                End If
                
                If k = UBound(AddressList) Then Exit Do
                
                k = k + 1
            
            Loop Until SearchFlag(j)
       
        Next j
  
        For j = 0 To UBound(SearchWord)
        
            If SearchFlag(j) Then
            
                ReDim Preserve SearchList(UBound(CityList, 1) + 1, SearchListCount)
                
                For k = 0 To UBound(CityList, 1)
                
                    SearchList(k, SearchListCount) = CityList(k, i)
                
                Next k
                
                SearchList(SearchListNo("SearchWord"), SearchListCount) = SearchWord(j)
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
        
                '�������[�h���s�撬�����ŏd���s�撬�����ɊY������ꍇ�͒���Ō���
                Case SearchList(SearchListNo("SearchWord"), i) = SearchList(SearchListNo("CityName"), i) _
                        And DupCityDic.Exists(SearchList(SearchListNo("PrefAreaCityName"), i))
                
                    For j = 0 To UBound(TownList, 2)
                    
                        If TownList(TownListNo("CityName"), j) = SearchList(SearchListNo("CityName"), i) Then
                        
                            SearchTownName = TownList(TownListNo("CityName"), j) & "*" & TownList(TownListNo("TownName"), j)
                            
                            Select Case True
                                    
                                '�����̒��悪����ꍇ
                                Case DupTownDic.Exists(SearchTownName)
                                
                                    If WorksheetFunction.CountIf(.Columns("A:A"), SearchTownName & "*") > 0 Then
                                    
                                        .Range("A1").AutoFilter Field:=1, Criteria1:=SearchTownName & "*"
                                        Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
                                        OutputRange.SpecialCells(xlCellTypeVisible).Value = "�������̒��悪���邽�ߗv�m�F��"
                                        Set OutputRange = Nothing
                                        .Range("A1").AutoFilter
                                        
                                    End If
                                
                                Case Else
                                
                                    If WorksheetFunction.CountIf(.Columns("A:A"), SearchTownName & "*") > 0 Then
                                    
                                        .Range("A1").AutoFilter Field:=1, Criteria1:=SearchTownName & "*"
                                        
                                        For k = 0 To 3
                                        
                                            Set OutputRange = .Range(.Cells(2, k + 2), .Cells(endRow, k + 2))
                                            OutputRange.SpecialCells(xlCellTypeVisible).Value = TownList(k, j)
                                            Set OutputRange = Nothing
                                        
                                        Next k
                                        
                                        .Range("A1").AutoFilter
                                        
                                    End If
                                    
                            End Select
                            
                        End If
                    
                    Next j
                
                Case Else
                
                    .Range("A1").AutoFilter Field:=1, Criteria1:=SearchList(UBound(SearchList, 1), i) & "*"
                    
                    For j = 0 To 3
                    
                        Set OutputRange = .Range(.Cells(2, j + 2), .Cells(endRow, j + 2))
                        OutputRange.SpecialCells(xlCellTypeVisible).Value = SearchList(j, i)
                        Set OutputRange = Nothing
                        
                    Next j
                    
                    .Range("A1").AutoFilter
                
            End Select
        
            
        End With
        
        Application.StatusBar = i & "/" & UBound(SearchList, 2) & "���ڂ��������ł�..."
        
    
    Next i
    

    Application.StatusBar = "���X�g���������ł�..."
   
    With TargetSheet
    
    
      '-------  �뎚�i�P�̃��j�m�F -------
    
        For i = 0 To TypoDic.Count - 1
            
            If WorksheetFunction.CountIf(.Range(.Range("A2"), .Cells(endRow, "A")), "*" & TypoDic.Items()(i)("TypoName") & "*") > 0 Then
                    
                .Range("A1").AutoFilter Field:=1, Criteria1:="*" & TypoDic.Items()(i)("TypoName") & "*"
                Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
                OutputRange.SpecialCells(xlCellTypeVisible).Value = "��" & TypoDic.Items()(i)("CorrectName") & "���������̂ł�" & "��"
                Set OutputRange = Nothing
                .Range("A1").AutoFilter
                
            End If
            
            
        Next i
    
    
         
        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:=""
            Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
            OutputRange.SpecialCells(xlCellTypeVisible).Value = "���Y������f�[�^���Ȃ����ߗv�m�F��"
            Set OutputRange = Nothing
            .Range("A1").AutoFilter
            
        End If
        
        
        '------- �v�m�F�Z���Ƀ}�[�J�[ -------
        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "��*") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:="��*"
            Set OutputRange = .Range(.Range("A2"), .Cells(endRow, "A"))
            OutputRange.SpecialCells(xlCellTypeVisible).Interior.Color = vbYellow
            Set OutputRange = Nothing
            .Range("A1").AutoFilter
            
        End If
        
        
        
        
        '�Z���⊮
        Call ComplementAddress(TargetSheet)
    
    
        .Range("A1").AutoFilter
    
    End With
    
    
    Set CityListNo = Nothing
    Set TownListNo = Nothing
    Set DupTownDic = Nothing
    Set DupCityDic = Nothing
    Set TypoDic = Nothing
    Set SearchListNo = Nothing
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
