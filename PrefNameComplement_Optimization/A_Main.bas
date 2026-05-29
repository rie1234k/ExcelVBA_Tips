Attribute VB_Name = "A_Main"
Option Explicit

Public Sub PrefNameComplement()


'--- URL・ファイル ---
Dim TargetURL As String
Dim PostcodeFileName As String

'--- シート処理 ---
Dim TargetSheet As Worksheet
Dim endRow As Long
Dim AddressData As Variant
Dim ResultData As Variant

'--- 検索ワード作成用 ---
Dim MaxCharCount As Long
Dim AddressCharCount As Long
Dim AddressIndex As Object
Dim SearchWordCheckIndex As Object
Dim TargetTownIndex As Object
Dim addrKey As Variant
Dim cityKey As Variant
Dim townKey As Variant
Dim typoKey As Variant
Dim SearKey As Variant
Dim TargetIndexString As String
Dim TargetChar As String
Dim DicTownWord As String

'--- データ ---
Dim ListCollection As Collection
Dim CityIndex As Object
Dim PrefAreaCityIndex As Object
Dim AreaCityIndex As Object
Dim PrefCityIndex As Object
Dim TownIndex As Object
Dim CityRec As Collection
Dim TownRec As Collection


'--- 重複・誤字管理 ---
Dim DupCityIndex As Object
Dim DupTownIndex As Object
Dim TypoIndex As Object

'--- 検索用 ---
Dim SearchWords As Object
Dim DicSearchName As String
Dim SearchPattern As String
Dim TargetAddress As String


'--- 状態フラグ ---
Dim isTargetWord As Boolean
Dim isDupCity As Boolean

'--- ループ ---
Dim i As Long
Dim j As Long


'--- 定数 ---
Const COL_CODE As Long = 1
Const COL_PREF As Long = 2
Const COL_AREA As Long = 3
Const COL_CITY As Long = 4


    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '郵便番号データダウンロード＞住所の郵便番号（CSV形式）＞読み仮名データの促音・拗音を小書きで表記するもの＞全国一括
    TargetURL = "https://www.post.japanpost.jp/service/search/zipcode/download/kogaki/zip/ken_all.zip"
    
    'ダウンロードしたZipファイル内のCSVファイル名
    PostcodeFileName = "KEN_ALL.CSV"
    
     Application.StatusBar = "データをダウンロード中です..."
    
    Call PostcodeZipFileDownload(TargetURL)
    
    Application.StatusBar = "市区町村リスト・町域リストを作成中です..."
    
    Set ListCollection = New Collection
    Set ListCollection = GetListCollection(PostcodeFileName)
    
    Set PrefAreaCityIndex = ListCollection("PrefAreaCityIndex")
    Set PrefCityIndex = ListCollection("PrefCityIndex")
    Set AreaCityIndex = ListCollection("AreaCityIndex")
    Set CityIndex = ListCollection("CityIndex")
    Set DupCityIndex = ListCollection("DupCityIndex")
    Set TownIndex = ListCollection("TownIndex")
    Set DupTownIndex = ListCollection("DupTownIndex")
    Set TypoIndex = ListCollection("TypoIndex")
    
    Set ListCollection = Nothing

    Application.StatusBar = "検索用ワードリストを作成中です..."
    Set TargetSheet = ThisWorkbook.Sheets("作業シート")

    With TargetSheet
        
        endRow = TargetSheet.Cells(Rows.Count, "A").End(xlUp).Row
        
        'オートフィルターが設定されている場合には解除
        If Not .AutoFilter Is Nothing Then .Range("A1").AutoFilter
        '塗りつぶし解除
        .Range("A1").CurrentRegion.Interior.Color = xlNone

        .Range("A1").CurrentRegion.Offset(1, 1).ClearContents
    
    End With
   '------- 検索対象ワードのうち最大文字数を取得 -------
    
    MaxCharCount = 0
    For Each cityKey In PrefAreaCityIndex.Keys
        If MaxCharCount < Len(cityKey) Then
            MaxCharCount = Len(cityKey)
        End If
    Next cityKey
    
    '------- 住所インデックスを作成(検索高速化用) -------
    Set AddressIndex = CreateObject("Scripting.Dictionary")
    Set SearchWordCheckIndex = CreateObject("Scripting.Dictionary")
    Set TargetTownIndex = CreateObject("Scripting.Dictionary")
    
    AddressCharCount = 0
    For i = 2 To endRow
        If MaxCharCount > 0 Then
            TargetIndexString = Left(TargetSheet.Cells(i, "A").Value, MaxCharCount)
        Else
            TargetIndexString = TargetSheet.Cells(i, "A").Value
            If AddressCharCount < Len(TargetIndexString) Then
                AddressCharCount = Len(TargetIndexString)
            End If
        End If
        
        If Not AddressIndex.Exists(TargetIndexString) Then
            AddressIndex.Add TargetIndexString, True
        End If
    Next i
    
    If AddressCharCount > 0 Then MaxCharCount = AddressCharCount
    
    Set SearchWords = CreateObject("Scripting.Dictionary")
   
   
     '------- 検索対象となる市区町村候補を生成 -------
    For Each addrKey In AddressIndex.Keys
            
        Set CityRec = Nothing
       
        For i = MaxCharCount To 1 Step -1
            TargetChar = Left(addrKey, i)
            If TownIndex.Exists(TargetChar) Then
                DicTownWord = TownIndex(TargetChar)("PrefAreaCityName") & "/" & TownIndex(TargetChar)("TownName")
                If Not TargetTownIndex.Exists(DicTownWord) Then
                    TargetTownIndex.Add DicTownWord, TownIndex(TargetChar)
                End If
            End If
            
            If PrefAreaCityIndex.Exists(TargetChar) Then Set CityRec = PrefAreaCityIndex(TargetChar): Exit For
            If PrefCityIndex.Exists(TargetChar) Then Set CityRec = PrefCityIndex(TargetChar): Exit For
            If AreaCityIndex.Exists(TargetChar) Then Set CityRec = AreaCityIndex(TargetChar): Exit For
            If CityIndex.Exists(TargetChar) Then Set CityRec = CityIndex(TargetChar): Exit For
        Next i
        
        If Not CityRec Is Nothing Then
            If Not SearchWordCheckIndex.Exists(TargetChar) Then
                SearchWordCheckIndex.Add TargetChar, True
                SearchWords.Add TargetChar, CityRec
            End If
        End If
    
    Next addrKey
    

    
    Set CityRec = Nothing
    Set SearchWordCheckIndex = Nothing
    Set PrefAreaCityIndex = Nothing
    Set CityIndex = Nothing
    Set TownIndex = Nothing
    Set PrefCityIndex = Nothing
    Set AreaCityIndex = Nothing

    '------- 都道府県・市区町村名取得処理 -------
        
    With TargetSheet
        
        AddressData = .Range(.Range("A2"), .Cells(endRow, "A")).Value
        ReDim ResultData(1 To endRow - 1, 1 To 4)
        
        For Each SearKey In SearchWords.Keys
            
            '検索ワードが市区町村名で重複市区町村名に該当する場合は町域で検索
            isTargetWord = SearKey = SearchWords(SearKey)("CityName")
            isDupCity = DupCityIndex.Exists(SearchWords(SearKey)("PrefAreaCityName"))
            
            If isTargetWord And isDupCity Then
               
               For Each townKey In TargetTownIndex
                 Set TownRec = TargetTownIndex(townKey)
                 If TownRec("CityName") = SearchWords(SearKey)("CityName") Then
                    
                    SearchPattern = TownRec("CityName") & "*" & TownRec("TownName") & "*"
                    DicSearchName = TownRec("CityName") & "/" & TownRec("TownName")
                     
                    '同名の町域がある場合
                    If DupTownIndex.Exists(DicSearchName) Then
                        For j = 1 To UBound(AddressData, 1)
                            TargetAddress = AddressData(j, 1)
                            If TargetAddress Like SearchPattern Then
                                ResultData(j, 1) = "★" & CreateDupTownMessage(DicSearchName, DupTownIndex) & "に同名の町域「" & DupTownIndex(DicSearchName)(1)("TownName") & "」があるため要確認★"
                            End If
                        Next j
                        
                       
                    '同名の町域がない場合
                    Else
                        For j = 1 To UBound(AddressData, 1)
                            TargetAddress = AddressData(j, 1)
                            If TargetAddress Like SearchPattern Then
                                ResultData(j, COL_CODE) = TownRec("CityCode")
                                ResultData(j, COL_PREF) = TownRec("PrefName")
                                ResultData(j, COL_AREA) = TownRec("AreaName")
                                ResultData(j, COL_CITY) = TownRec("CityName")
                            End If
                        Next j
                    End If
                 End If
               Next townKey
             
            '重複市区町村名ではない場合
            Else
                SearchPattern = SearKey & "*"
                For j = 1 To UBound(AddressData, 1)
                    TargetAddress = AddressData(j, 1)
                    If TargetAddress Like SearchPattern Then
                        ResultData(j, COL_CODE) = SearchWords(SearKey)("CityCode")
                        ResultData(j, COL_PREF) = SearchWords(SearKey)("PrefName")
                        ResultData(j, COL_AREA) = SearchWords(SearKey)("AreaName")
                        ResultData(j, COL_CITY) = SearchWords(SearKey)("CityName")
                    End If
                Next j
             
            End If
        
        Next SearKey

    
      '-------  誤字（ケ⇔ヶ）確認 -------
        
        For Each typoKey In TypoIndex
            SearchPattern = "*" & TypoIndex.Item(typoKey)("TypoName") & "*"
            For j = 1 To UBound(AddressData, 1)
                TargetAddress = AddressData(j, 1)
                If TargetAddress Like SearchPattern Then
                    ResultData(j, 1) = "★" & TypoIndex.Item(typoKey)("CorrectName") & "が正式名称のため要確認" & "★"
                End If
            Next j
        Next typoKey
        
               
      '------- 不明の住所確認 -------
        For j = 1 To UBound(AddressData, 1)
            If ResultData(j, 1) = "" Then
                ResultData(j, 1) = "★該当するデータがないため要確認★"
            End If
        Next j

        .Range(.Range("B2"), .Cells(endRow, "E")).Value = ResultData
      
      '------- 要確認住所にマーカー -------
        
        For j = 1 To UBound(ResultData, 1)
            
            If ResultData(j, 1) Like "★*" Then
                .Cells(j + 1, "A").Interior.Color = vbYellow
            End If
 
        Next j
      
        '住所補完
        Call ComplementAddress(TargetSheet)

        .Range("A1").AutoFilter
    
    End With
    
Cleanup:
    
    Set DupTownIndex = Nothing
    Set DupCityIndex = Nothing
    Set TargetTownIndex = Nothing
    Set TypoIndex = Nothing
    Set TargetSheet = Nothing
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrHandler:

    MsgBox "発生元：" & Err.Source & vbCrLf & _
           "内容　：" & Err.Description
    
    GoTo Cleanup

End Sub

Public Function CreateDupTownMessage(DicSearchName As String, DupTownIndex As Object) As String
Dim i As Long

    CreateDupTownMessage = ""
    For i = 1 To DupTownIndex(DicSearchName).Count
        If CreateDupTownMessage = "" Then
            CreateDupTownMessage = DupTownIndex(DicSearchName)(i)("PrefAreaCityName")
        Else
            CreateDupTownMessage = CreateDupTownMessage & "と" & DupTownIndex(DicSearchName)(i)("PrefAreaCityName")
        End If
    Next i

End Function


Public Sub ComplementAddress(TargetSheet As Worksheet)

Dim endRow As Long
Dim ReplaceAddress As String
Dim FullAddress As String
Dim i As Long
Dim j As Long
Dim TargetData As Variant
Dim Result As Variant

    With TargetSheet
    
        endRow = .Cells(Rows.Count, "A").End(xlUp).Row
        
        ReDim Result(1 To endRow - 1, 1 To 2)
        TargetData = .Range(.Range("A2"), .Cells(endRow, "E")).Value
        
        For i = 1 To UBound(TargetData, 1)
            If TargetData(i, 3) <> "" Then
                ReplaceAddress = TargetData(i, 1)
                FullAddress = ""
                For j = 3 To 5
                    ReplaceAddress = Replace(ReplaceAddress, TargetData(i, j), "")
                    FullAddress = FullAddress & TargetData(i, j)
                Next j
                Result(i, 1) = ReplaceAddress
                Result(i, 2) = FullAddress & ReplaceAddress
            End If
        Next i

        .Range(.Range("F2"), .Cells(endRow, "G")).Value = Result

    End With

End Sub


