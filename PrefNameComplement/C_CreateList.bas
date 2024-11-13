Attribute VB_Name = "C_CreateList"
Option Explicit

'郵便番号データファイルの列番号
Enum CSVColumnNo

    CityCode = 1
    PrefName = 7
    CityName = 8
    TownName = 9
 
End Enum

Private CityList() As String
Private TownList() As String
Private DupTownDic As Object
Private DupCityDic As Object

Public Function GetListCollection(PostcodeFileName As String) As Collection

    Call CreateCityList(PostcodeFileName)
    Call CreateTownList(PostcodeFileName)
    
    Set GetListCollection = New Collection
    
    GetListCollection.Add CityList, "CityList"
    GetListCollection.Add TownList, "TownList"
    GetListCollection.Add DupTownDic, "DupTownDic"
    GetListCollection.Add DupCityDic, "DupCityDic"
    
End Function
Public Sub CreateCityList(PostcodeFileName As String)

Dim FSO As Object
Dim TextFile As Object
Dim SplitData As Variant
Dim i As Long

Dim CityListSet As Collection
Dim CityCode As String
Dim PrefName As String
Dim BaseCityName As String
Dim CityName As String
Dim AreaName As String
Dim BeforeFullCityName As String
Dim CityCount As Long
Dim CityDic As Object
 
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TextFile = FSO.OpenTextFile(ThisWorkbook.Path & "\" & PostcodeFileName)
   
    CityCount = 0
   
    Set CityDic = CreateObject("Scripting.Dictionary")
    Set DupCityDic = CreateObject("Scripting.Dictionary")
    
    Do
        
        SplitData = Split(TextFile.ReadLine, ",")
        
        CityCode = Replace(SplitData(CSVColumnNo.CityCode - 1), """", "")
        PrefName = Replace(SplitData(CSVColumnNo.PrefName - 1), """", "")
        BaseCityName = Replace(SplitData(CSVColumnNo.CityName - 1), """", "")
        
        '------- CSVデータ加工 -------
        Select Case Val(Left(Right(CityCode, 3), 1))
        
            Case 1 '全国地方公共団体コード 後ろから3つ目が100～199⇒政令指定都市(行政区あり)
            
                AreaName = ""
                
                '東京都特別区はそのまま
                If PrefName = "東京都" Then
                
                    CityName = BaseCityName
                    
                Else
                
                    '市名の抜き出し
                    CityName = Left(BaseCityName, InStr(BaseCityName, "市"))
                    
                    '行政区の地方公共団体コードを市の地方公共団体コードに変換
                    If Val(Left(Right(CityCode, 3), 1)) = 1 Then
                    
                        CityCode = Left(CityCode, Len(CityCode) - 1) & "0"
                    
                    End If
                
                End If
            
            Case 2 '全国地方公共団体コード 後ろから3つ目が200～299⇒市（政令指定都市以外）
            
                AreaName = ""
                CityName = BaseCityName
                
            Case Is >= 3 '全国地方公共団体コード 後ろから3つ目が300以上⇒町村（郡あり）
            
                AreaName = Left(BaseCityName, InStr(BaseCityName, "郡"))
                CityName = Replace(BaseCityName, AreaName, "")
            
        End Select
        
        
        Set CityListSet = New Collection
        
        CityListSet.Add CityCode, "CityCode"
        CityListSet.Add PrefName, "PrefName"
        CityListSet.Add AreaName, "AreaName"
        CityListSet.Add CityName, "CityName"
        CityListSet.Add PrefName & AreaName & CityName, "FullCityName"
        CityListSet.Add AreaName & CityName, "FullAreaName"
        
        '------- 市区町村リスト作成 -------
        If CityListSet("FullCityName") <> BeforeFullCityName Then
        
            ReDim Preserve CityList(5, CityCount)
            
            For i = 0 To UBound(CityList, 1)
            
                CityList(i, CityCount) = CityListSet(i + 1)
            
            Next i
            
            
            '------- 重複確認用辞書作成 -------
            If Not CityDic.Exists(CityName) Then
            
                CityDic.Add CityName, CityListSet
            
            Else
            
                If Not DupCityDic.Exists(CityDic(CityName)("FullCityName")) Then
                
                DupCityDic.Add CityDic(CityName)("FullCityName"), CityDic(CityName)
                
                End If
                
                DupCityDic.Add CityListSet("FullCityName"), CityListSet
            
            End If
            
            BeforeFullCityName = CityListSet("FullCityName")
            CityCount = CityCount + 1
            Set CityListSet = Nothing
        
        End If
        
    Loop Until TextFile.AtEndOfLine
    
    Set CityDic = Nothing
    Set TextFile = Nothing
    Set FSO = Nothing
   
    
End Sub

Public Sub CreateTownList(PostcodeFileName As String)

Dim FSO As Object
Dim TextFile As Object
Dim SplitData As Variant
Dim i As Long

Dim TownDic As Object
Dim PrefName As String
Dim FullAreaName As String
Dim CityName As String
Dim TownName As String

Dim TownCount As Long
Dim DupTownCount As Long

   
    '------- 検索用重複市区町村町域リスト作成 -------
    TownCount = 0
    DupTownCount = 0
    
    Set TownDic = CreateObject("Scripting.Dictionary")
    Set DupTownDic = CreateObject("Scripting.Dictionary")

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TextFile = FSO.OpenTextFile(ThisWorkbook.Path & "\" & PostcodeFileName)
   
    Do
    
        SplitData = Split(TextFile.ReadLine, ",")

        PrefName = Replace(SplitData(CSVColumnNo.PrefName - 1), """", "")
        FullAreaName = Replace(SplitData(CSVColumnNo.CityName - 1), """", "")
        TownName = Replace(SplitData(CSVColumnNo.TownName - 1), """", "")
        
        If DupCityDic.Exists(PrefName & FullAreaName) Then
        
           CityName = DupCityDic(PrefName & FullAreaName)("CityName")
             
             If TownName <> "以下に掲載がない場合" Then
             
                If Not TownDic.Exists(CityName & TownName) Then
                
                    ReDim Preserve TownList(4, TownCount)
                    
                    For i = 0 To 3
                    
                        TownList(i, TownCount) = DupCityDic(PrefName & FullAreaName)(i + 1)
                        
                    Next i

                    TownList(4, TownCount) = TownName
                    
                    TownDic.Add CityName & TownName, CityName & TownName
                    
                    TownCount = TownCount + 1
                
                Else
                
                    DupTownDic.Add CityName & "*" & TownName, CityName & "*" & TownName
                    DupTownCount = DupTownCount + 1
                    
                End If
                 
             End If
             
        End If
 
     Loop Until TextFile.AtEndOfLine
     
    Set TextFile = Nothing
    Set FSO = Nothing

 
End Sub

