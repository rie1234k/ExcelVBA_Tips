Attribute VB_Name = "C_CreateIndex"
Option Explicit

'郵便番号データファイルの列番号
Private Enum CSVColumnNo

    CityCode = 1
    PrefName = 7
    CityName = 8
    TownName = 9
 
End Enum

Private PrefAreaCityIndex As Object
Private PrefCityIndex As Object
Private AreaCityIndex As Object
Private CityIndex As Object
Private DupCityIndex As Object
Private TownIndex As Object
Private DupTownIndex As Object
Private TypoIndex As Object


Public Function GetListCollection(PostcodeFileName As String) As Collection
        
    Call CreateCityIndex(PostcodeFileName)
    Call CreateTownIndex(PostcodeFileName)
    
    Set GetListCollection = New Collection
    
    GetListCollection.Add PrefAreaCityIndex, "PrefAreaCityIndex"
    GetListCollection.Add PrefCityIndex, "PrefCityIndex"
    GetListCollection.Add AreaCityIndex, "AreaCityIndex"
    GetListCollection.Add CityIndex, "CityIndex"
    GetListCollection.Add DupCityIndex, "DupCityIndex"
    GetListCollection.Add TownIndex, "TownIndex"
    GetListCollection.Add DupTownIndex, "DupTownIndex"
    GetListCollection.Add TypoIndex, "TypoIndex"
   
End Function


Private Sub CreateCityIndex(PostcodeFileName As String)

Dim Fso As Object
Dim TextFile As Object
Dim SplitData As Variant
Dim i As Long

Dim CityRec As Collection
Dim CityCode As String
Dim PrefName As String
Dim BaseCityName As String
Dim CityName As String
Dim AreaName As String
Dim TypoName As String
Dim AddItem As Collection
Dim BeforePrefAreaCityName As String
Dim codeType As Long
Dim CityCheckIndex As Object

 
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Set TextFile = Fso.OpenTextFile(ThisWorkbook.Path & "\" & PostcodeFileName)
   
    Set PrefAreaCityIndex = CreateObject("Scripting.Dictionary")
    Set PrefCityIndex = CreateObject("Scripting.Dictionary")
    Set AreaCityIndex = CreateObject("Scripting.Dictionary")
    Set CityIndex = CreateObject("Scripting.Dictionary")
    Set CityCheckIndex = CreateObject("Scripting.Dictionary")
    Set DupCityIndex = CreateObject("Scripting.Dictionary")
    Set TypoIndex = CreateObject("Scripting.Dictionary")
    
    Do
        
        SplitData = Split(TextFile.ReadLine, ",")
        
        CityCode = Replace(SplitData(CSVColumnNo.CityCode - 1), """", "")
        PrefName = Replace(SplitData(CSVColumnNo.PrefName - 1), """", "")
        BaseCityName = Replace(SplitData(CSVColumnNo.CityName - 1), """", "")
        
        CityName = NormalizeCity(CityCode, PrefName, BaseCityName, AreaName)
        
        '政令指定都市の場合は区コード→市コードへ変換
        If (CLng(CityCode) Mod 1000) \ 100 = 1 Then
            CityCode = Format((CLng(CityCode) \ 100) * 100, "00000")
        End If
        
        Set CityRec = New Collection
        
        CityRec.Add CityCode, "CityCode"
        CityRec.Add PrefName, "PrefName"
        CityRec.Add AreaName, "AreaName"
        CityRec.Add CityName, "CityName"
        CityRec.Add PrefName & AreaName & CityName, "PrefAreaCityName"
        CityRec.Add AreaName & CityName, "AreaCityName"
        CityRec.Add PrefName & CityName, "PrefCityName"
        
        '------- 検索用辞書作成 -------
        If CityRec("PrefAreaCityName") <> BeforePrefAreaCityName Then
                
            If Not PrefAreaCityIndex.Exists(CityRec("PrefAreaCityName")) Then
                    PrefAreaCityIndex.Add CityRec("PrefAreaCityName"), CityRec
            End If
                
            If Not PrefCityIndex.Exists(CityRec("PrefCityName")) Then
                    PrefCityIndex.Add CityRec("PrefCityName"), CityRec
            End If
            
            If Not AreaCityIndex.Exists(CityRec("AreaCityName")) Then
                    AreaCityIndex.Add CityRec("AreaCityName"), CityRec
            End If
                                
            If Not CityIndex.Exists(CityRec("CityName")) Then
                    CityIndex.Add CityRec("CityName"), CityRec
            End If
              
            '------- 重複確認用辞書作成 -------
            If Not CityCheckIndex.Exists(CityName) Then
                        
                CityCheckIndex.Add CityName, CityRec
                
            Else
            
                If Not DupCityIndex.Exists(CityCheckIndex(CityName)("PrefAreaCityName")) Then
                    DupCityIndex.Add CityCheckIndex(CityName)("PrefAreaCityName"), CityIndex(CityName)
                End If
                
                If Not DupCityIndex.Exists(CityRec("PrefAreaCityName")) Then
                    DupCityIndex.Add CityRec("PrefAreaCityName"), CityRec
                End If
                
            End If
            
            '------- 誤字（ケ⇔ヶ）確認用辞書作成 -------
            For i = 2 To 4

                Select Case True
                
                    Case CityRec(i) Like "*ケ*"
                    
                        TypoName = Replace(CityRec(i), "ケ", "ヶ")
                        
                    Case CityRec(i) Like "*ヶ*"
                
                        TypoName = Replace(CityRec(i), "ヶ", "ケ")
                        
                    Case Else
                    
                        TypoName = ""

                End Select
                
                If TypoName <> "" Then
                    
                    If TypoIndex.Exists(CityRec(i)) Then
                        
                        TypoIndex.Remove CityRec(i)
                    
                    Else
                    
                        If Not TypoIndex.Exists(TypoName) Then
                            
                            Set AddItem = New Collection
                            
                            AddItem.Add CityRec(i), "CorrectName"
                            AddItem.Add TypoName, "TypoName"
                            TypoIndex.Add TypoName, AddItem
                            
                            Set AddItem = Nothing
                        
                        End If
                        
                    End If
                    
                End If
                
            Next i

            BeforePrefAreaCityName = CityRec("PrefAreaCityName")
            Set CityRec = Nothing
    
        End If
        
    Loop Until TextFile.AtEndOfStream
    
    TextFile.Close

    Set TextFile = Nothing
    Set Fso = Nothing
   
    
End Sub

Private Sub CreateTownIndex(PostcodeFileName As String)

Dim Fso As Object
Dim TextFile As Object
Dim SplitData As Variant
Dim i As Long
Dim CityCode As String
Dim PrefName As String
Dim BaseCityName As String
Dim CityName As String
Dim AreaName As String
Dim TownName As String
Dim TownRec As Collection
Dim DupTownRec As Collection
Dim TownCheckIndex As Object
Dim townKey As String
 
    '------- 検索用重複市区町村町域リスト作成 -------
    
    Set TownIndex = CreateObject("Scripting.Dictionary")
    Set TownCheckIndex = CreateObject("Scripting.Dictionary")
    Set DupTownIndex = CreateObject("Scripting.Dictionary")
   
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    
    Set TextFile = Fso.OpenTextFile(ThisWorkbook.Path & "\" & PostcodeFileName)
    
    
    Do
    
        SplitData = Split(TextFile.ReadLine, ",")
        CityCode = Replace(SplitData(CSVColumnNo.CityCode - 1), """", "")
        PrefName = Replace(SplitData(CSVColumnNo.PrefName - 1), """", "")
        BaseCityName = Replace(SplitData(CSVColumnNo.CityName - 1), """", "")
        
        CityName = NormalizeCity(CityCode, PrefName, BaseCityName, AreaName)
        
        If DupCityIndex.Exists(PrefName & AreaName & CityName) Then
            
           TownName = Replace(SplitData(CSVColumnNo.TownName - 1), """", "")
          
             If TownName <> "以下に掲載がない場合" Then

                Set TownRec = New Collection
                
                TownRec.Add DupCityIndex(PrefName & AreaName & CityName)("CityCode"), "CityCode"
                TownRec.Add PrefName, "PrefName"
                TownRec.Add DupCityIndex(PrefName & AreaName & CityName)("AreaName"), "AreaName"
                TownRec.Add DupCityIndex(PrefName & AreaName & CityName)("CityName"), "CityName"
                TownRec.Add TownName, "TownName"
                TownRec.Add PrefName & AreaName & CityName & TownName, "PrefAreaCityTownName"
                TownRec.Add PrefName & AreaName & CityName, "PrefAreaCityName"
                TownRec.Add CityName & TownName, "CityTownName"
                
                townKey = CityName & "/" & TownName
                
                If Not TownCheckIndex.Exists(townKey) Then
                    
                    TownCheckIndex.Add townKey, TownRec
                    TownIndex.Add CityName & TownName, TownRec '検索用辞書
                    TownIndex.Add CityName & "大字" & TownName, TownRec
                Else
                     
                    Set DupTownRec = New Collection
                    
                    If Not DupTownIndex.Exists(townKey) Then
                        DupTownRec.Add TownCheckIndex(townKey), TownCheckIndex(townKey)("PrefAreaCityTownName")
                        DupTownIndex.Add townKey, DupTownRec
                    End If
                    
                    DupTownIndex(townKey).Add TownRec, TownRec("PrefAreaCityTownName")
       
                End If
                
                Set TownRec = Nothing
                 
             End If
             
        End If
 
     Loop Until TextFile.AtEndOfStream

End Sub


Private Function NormalizeCity(ByVal CityCode As String, ByVal PrefName As String, ByVal BaseCityName As String, ByRef AreaName As String) As String

Dim codeType As Long
Dim pos As Long
    
    codeType = (CLng(CityCode) Mod 1000) \ 100
    
    
    Select Case codeType
    
        Case 1 '政令指定都市
            AreaName = ""
            
            If PrefName = "東京都" Then
                NormalizeCity = BaseCityName
            Else
                pos = InStr(BaseCityName, "市")
                If pos > 0 Then
                    NormalizeCity = Left(BaseCityName, pos)
                Else
                    NormalizeCity = BaseCityName
                End If
            End If
            
        Case 2 '市
            AreaName = ""
            NormalizeCity = BaseCityName
        Case Else '町村
            pos = InStr(BaseCityName, "郡")
            If pos > 0 Then
                AreaName = Left(BaseCityName, pos)
                NormalizeCity = Mid(BaseCityName, pos + 1)
            Else
                AreaName = ""
                NormalizeCity = BaseCityName
            End If
    End Select

End Function


