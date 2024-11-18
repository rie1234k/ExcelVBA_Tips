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


    '郵便番号データダウンロード＞住所の郵便番号（CSV形式）＞読み仮名データの促音・拗音を小書きで表記するもの＞全国一括
    TargetUrl = "https://www.post.japanpost.jp/zipcode/dl/kogaki/zip/ken_all.zip"
    
    'ダウンロードしたZipファイル内のCSVファイル名
    PostcodeFileName = "KEN_ALL.CSV"
    
     Application.StatusBar = "データをダウンロード中です..."
    
    Call PostcodeZipFileDowunload(TargetUrl, PostcodeFileName)
 
 
    Application.StatusBar = "市区町村リスト・町域リストを作成中です..."
    
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

    Application.StatusBar = "検索用ワードリストを作成中です..."
    
    Set TargetSheet = ThisWorkbook.Sheets("作業シート")
    endRow = TargetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    
     '------- 対象住所リストを配列として取得 -------
    ReDim AddressList(2 To endRow)
    
    For i = 2 To endRow
        
        AddressList(i) = TargetSheet.Cells(i, "A").Value

    Next i
    

    '------- 検索用ワードリスト作成 -------
    
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
    
    '------- 都道府県・市区町村名取得処理 -------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
     '塗りつぶし解除
    TargetSheet.Range("A1").CurrentRegion.Interior.Color = xlNone

    'オートフィルターが設定されている場合には解除
    If Not TargetSheet.AutoFilter Is Nothing Then TargetSheet.Range("A1").AutoFilter
    
    TargetSheet.Range("A1").CurrentRegion.Offset(1, 1).ClearContents
    
    For i = 0 To UBound(SearchList, 2)
        
        With TargetSheet
            
            Select Case True
        
                '検索ワードが市区町村名で重複市区町村名に該当する場合は町域で検索
                Case SearchList(SearchListNo("SearchWord"), i) = SearchList(SearchListNo("CityName"), i) _
                        And DupCityDic.Exists(SearchList(SearchListNo("PrefAreaCityName"), i))
                
                    For j = 0 To UBound(TownList, 2)
                    
                        If TownList(TownListNo("CityName"), j) = SearchList(SearchListNo("CityName"), i) Then
                        
                            SearchTownName = TownList(TownListNo("CityName"), j) & "*" & TownList(TownListNo("TownName"), j)
                            
                            Select Case True
                                    
                                '同名の町域がある場合
                                Case DupTownDic.Exists(SearchTownName)
                                
                                    If WorksheetFunction.CountIf(.Columns("A:A"), SearchTownName & "*") > 0 Then
                                    
                                        .Range("A1").AutoFilter Field:=1, Criteria1:=SearchTownName & "*"
                                        Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
                                        OutputRange.SpecialCells(xlCellTypeVisible).Value = "★同名の町域があるため要確認★"
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
        
        Application.StatusBar = i & "/" & UBound(SearchList, 2) & "件目を検索中です..."
        
    
    Next i
    

    Application.StatusBar = "リストを処理中です..."
   
    With TargetSheet
    
    
      '-------  誤字（ケ⇔ヶ）確認 -------
    
        For i = 0 To TypoDic.Count - 1
            
            If WorksheetFunction.CountIf(.Range(.Range("A2"), .Cells(endRow, "A")), "*" & TypoDic.Items()(i)("TypoName") & "*") > 0 Then
                    
                .Range("A1").AutoFilter Field:=1, Criteria1:="*" & TypoDic.Items()(i)("TypoName") & "*"
                Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
                OutputRange.SpecialCells(xlCellTypeVisible).Value = "★" & TypoDic.Items()(i)("CorrectName") & "が正式名称です" & "★"
                Set OutputRange = Nothing
                .Range("A1").AutoFilter
                
            End If
            
            
        Next i
    
    
         
        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:=""
            Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
            OutputRange.SpecialCells(xlCellTypeVisible).Value = "★該当するデータがないため要確認★"
            Set OutputRange = Nothing
            .Range("A1").AutoFilter
            
        End If
        
        
        '------- 要確認住所にマーカー -------
        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "★*") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:="★*"
            Set OutputRange = .Range(.Range("A2"), .Cells(endRow, "A"))
            OutputRange.SpecialCells(xlCellTypeVisible).Interior.Color = vbYellow
            Set OutputRange = Nothing
            .Range("A1").AutoFilter
            
        End If
        
        
        
        
        '住所補完
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
