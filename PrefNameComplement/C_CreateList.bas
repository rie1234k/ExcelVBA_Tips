Attribute VB_Name = "C_CreateList"
Option Explicit

'�X�֔ԍ��f�[�^�t�@�C���̗�ԍ�
Private Enum CSVColumnNo

    CityCode = 1
    PrefName = 7
    CityName = 8
    TownName = 9
 
End Enum

Private CityList() As String
Private TownList() As String
Private DupTownDic As Object
Private DupCityDic As Object
Private ListNo As Collection


Public Function GetListCollection(PostcodeFileName As String) As Collection

    Call SetListNo
    
    Call CreateCityList(PostcodeFileName)
    Call CreateTownList(PostcodeFileName)
    
    Set GetListCollection = New Collection
    
    GetListCollection.Add CityList, "CityList"
    GetListCollection.Add TownList, "TownList"
    GetListCollection.Add ListNo, "ListNo"
    GetListCollection.Add DupTownDic, "DupTownDic"
    GetListCollection.Add DupCityDic, "DupCityDic"
    
End Function
Private Sub SetListNo()

Dim CityList As Collection
Dim TownList As Collection
Dim i As Long
Dim ItemNameArray As Variant
    
    
    Set ListNo = New Collection

    'CityListNo�ݒ�
    Set CityList = New Collection
    ItemNameArray = Array("CityCode", "PrefName", "AreaName", "CityName", "PrefAreaCityName", "AreaCityName", "PrefCityName")
    
    For i = 0 To UBound(ItemNameArray)
         
        CityList.Add i, ItemNameArray(i)

    Next i

    ListNo.Add CityList, "CityList"
    
    Set CityList = Nothing
    
    
    'TownListNo�ݒ�
    Set TownList = New Collection
    ItemNameArray = Array("CityCode", "PrefName", "AreaName", "CityName", "TownName")
    
    For i = 0 To UBound(ItemNameArray)

        TownList.Add i, ItemNameArray(i)

    Next i
    
     ListNo.Add TownList, "TownList"
    
    Set TownList = Nothing
    

End Sub

Private Sub CreateCityList(PostcodeFileName As String)

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
Dim BeforePrefAreaCityName As String
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
        
         
        '------- CSV�f�[�^���H -------
        Select Case Val(Left(Right(CityCode, 3), 1))
        
            Case 1 '�S���n�������c�̃R�[�h ��납��3�ڂ�100�`199�ː��ߎw��s�s(�s���悠��)
            
                AreaName = ""
                
                '�����s���ʋ�͂��̂܂�
                If PrefName = "�����s" Then
                
                    CityName = BaseCityName
                    
                Else
                
                    '�s���̔����o��
                    CityName = Left(BaseCityName, InStr(BaseCityName, "�s"))
                    
                    '�s����̒n�������c�̃R�[�h���s�̒n�������c�̃R�[�h�ɕϊ�
                    If Val(Left(Right(CityCode, 3), 1)) = 1 Then
                    
                        CityCode = Left(CityCode, Len(CityCode) - 1) & "0"
                    
                    End If
                
                End If
            
            Case 2 '�S���n�������c�̃R�[�h ��납��3�ڂ�200�`299�ˎs�i���ߎw��s�s�ȊO�j
            
                AreaName = ""
                CityName = BaseCityName
                
            Case Is >= 3 '�S���n�������c�̃R�[�h ��납��3�ڂ�300�ȏ�˒����i�S����j
            
                AreaName = Left(BaseCityName, InStr(BaseCityName, "�S"))
                CityName = Replace(BaseCityName, AreaName, "")
            
        End Select
        
        
        Set CityListSet = New Collection
        
        CityListSet.Add CityCode, "CityCode"
        CityListSet.Add PrefName, "PrefName"
        CityListSet.Add AreaName, "AreaName"
        CityListSet.Add CityName, "CityName"
        CityListSet.Add PrefName & AreaName & CityName, "PrefAreaCityName"
        CityListSet.Add AreaName & CityName, "AreaCityName"
        CityListSet.Add PrefName & CityName, "PrefCityName"
            
        '------- �s�撬�����X�g�쐬 -------
        If CityListSet("PrefAreaCityName") <> BeforePrefAreaCityName Then
            
        
            ReDim Preserve CityList(CityListSet.Count - 1, CityCount)
            
            For i = 0 To UBound(CityList, 1)
            
                CityList(i, CityCount) = CityListSet(i + 1)
            
            Next i
            
            
            '------- �d���m�F�p�����쐬 -------
            If Not CityDic.Exists(CityName) Then
            
                CityDic.Add CityName, CityListSet
            
            Else
            
                If Not DupCityDic.Exists(CityDic(CityName)("PrefAreaCityName")) Then
                
                DupCityDic.Add CityDic(CityName)("PrefAreaCityName"), CityDic(CityName)
                
                End If
                
                DupCityDic.Add CityListSet("PrefAreaCityName"), CityListSet
            
            End If
            
            BeforePrefAreaCityName = CityListSet("PrefAreaCityName")
            CityCount = CityCount + 1
            Set CityListSet = Nothing
        
        End If
        
    Loop Until TextFile.AtEndOfLine
    
    Set CityDic = Nothing
    Set TextFile = Nothing
    Set FSO = Nothing
   
    
End Sub

Private Sub CreateTownList(PostcodeFileName As String)

Dim FSO As Object
Dim TextFile As Object
Dim SplitData As Variant
Dim i As Long

Dim TownDic As Object
Dim PrefName As String
Dim BaseCityName As String
Dim CityName As String
Dim TownName As String
Dim TownListSet As Collection


Dim TownCount As Long
Dim DupTownCount As Long

   
    '------- �����p�d���s�撬�����惊�X�g�쐬 -------
    TownCount = 0
    DupTownCount = 0
    
    Set TownDic = CreateObject("Scripting.Dictionary")
    Set DupTownDic = CreateObject("Scripting.Dictionary")

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TextFile = FSO.OpenTextFile(ThisWorkbook.Path & "\" & PostcodeFileName)
   
    Do
    
        SplitData = Split(TextFile.ReadLine, ",")

        PrefName = Replace(SplitData(CSVColumnNo.PrefName - 1), """", "")
        BaseCityName = Replace(SplitData(CSVColumnNo.CityName - 1), """", "")
        
        If DupCityDic.Exists(PrefName & BaseCityName) Then
            
           CityName = DupCityDic(PrefName & BaseCityName)("CityName")
           TownName = Replace(SplitData(CSVColumnNo.TownName - 1), """", "")
          
             If TownName <> "�ȉ��Ɍf�ڂ��Ȃ��ꍇ" Then
             
                If Not TownDic.Exists(CityName & TownName) Then
                    
                    Set TownListSet = New Collection
                    
                    TownListSet.Add DupCityDic(PrefName & BaseCityName)("CityCode"), "CityCode"
                    TownListSet.Add PrefName, "PrefName"
                    TownListSet.Add DupCityDic(PrefName & BaseCityName)("AreaName"), "AreaName"
                    TownListSet.Add DupCityDic(PrefName & BaseCityName)("CityName"), "CityName"
                    TownListSet.Add TownName, "TownName"

                    ReDim Preserve TownList(4, TownCount)
                    
                    For i = 0 To 4
                    
                        TownList(i, TownCount) = TownListSet(i + 1)
                        
                    Next i

                    TownDic.Add CityName & TownName, CityName & TownName
                    
                    TownCount = TownCount + 1
                    Set TownListSet = Nothing
                    
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

