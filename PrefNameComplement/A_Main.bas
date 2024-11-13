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
    Set DupTownDic = ListCollection("DupTownDic")
    Set DupCityDic = ListCollection("DupCityDic")
    
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
    For i = 0 To UBound(CityList, 2)
    
        SearchListWord(0) = CityList(3, i)
        SearchListWord(1) = CityList(4, i)
        SearchListWord(2) = CityList(5, i)
        
        '市町村名と郡+市町村名が同じ場合、市町村名のみの検索でよいので、空欄とする
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
                
                '検索ワードと市町村名が同じ時に重複市区町村名の場合は町域で検索
                Case SearchList(6, i) = SearchList(3, i) And DupCityDic.Exists(SearchList(4, i))
                    
                    For j = 0 To UBound(TownList, 2)
                        
                        If TownList(3, j) = SearchList(3, i) Then
                        
                            SearchTownName = TownList(3, j) & "*" & TownList(4, j)
                            
                            Select Case True
                                
                                '同名の町域がある場合
                                Case DupTownDic.Exists(SearchTownName)
                            
                                   If WorksheetFunction.CountIf(.Columns("A:A"), SearchTownName & "*") > 0 Then
    
                                    .Range("A1").AutoFilter Field:=1, Criteria1:=SearchTownName & "*"
                                    Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
                                    OutputRange.SpecialCells(xlCellTypeVisible).Value = "★同名の町域があるため要確認★"
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
        
        Application.StatusBar = i & "/" & UBound(SearchList, 2) & "件目を検索中です..."
        
    
    Next i
    

    Application.StatusBar = "リストを処理中です..."
    
    '------- 要確認住所にマーカー -------
    
    With TargetSheet

        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:=""
            Set OutputRange = .Range(.Range("B2"), .Cells(endRow, "B"))
            OutputRange.SpecialCells(xlCellTypeVisible).Value = "★該当するデータがないため要確認★"
        
        End If
        
        .Range("A1").AutoFilter
        
        If WorksheetFunction.CountIf(.Range(.Range("B2"), .Cells(endRow, "B")), "★*") > 0 Then
        
            .Range("A1").AutoFilter Field:=2, Criteria1:="★*"
            Set OutputRange = .Range(.Range("A2"), .Cells(endRow, "A"))
            OutputRange.SpecialCells(xlCellTypeVisible).Interior.Color = vbYellow
        
        End If
        
        .Range("A1").AutoFilter
        
        
        '住所補完
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
