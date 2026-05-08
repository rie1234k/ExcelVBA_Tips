Attribute VB_Name = "SearchSubfolders"
Option Explicit


Public Sub SearchSubFolders_File()

Dim Fso As Object
Dim FolderPath As String
Dim ChangePath As String
Dim StartRow As Long
Dim FolderStartColumn As Long
Dim TargetSheet As Worksheet

    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Set TargetSheet = ThisWorkbook.Sheets("ファイル名取得2")
    
    With TargetSheet
    
        .Range("A1").CurrentRegion.Offset(2).ClearContents
        
        FolderPath = .Range("B1").Value
        StartRow = .Range("A3").Row
        FolderStartColumn = .Range("C3").Column

    End With
    
    ChangePath = ChangeShortPath(FolderPath)
    
    If ChangePath <> "" Then
    
        Call FileSearch(TargetSheet, FolderPath, ChangePath, StartRow, FolderStartColumn, FolderStartColumn, Fso)
    
    Else
    
        MsgBox FolderPath & "は存在しません。"
        
    End If
    
    Set Fso = Nothing
    
End Sub

Sub FileSearch(TargetSheet As Worksheet, FolderPath As String, ChangePath As String, outRow As Long, outColumn As Long, baseColumn As Long, Fso As Object)

Dim i As Long
Dim TargetFolder As Object
Dim TargetSubfolder As Object
Dim TargetFile As Object
Dim CurrentFolderPath As String

    
    Set TargetFolder = Fso.GetFolder(ChangePath)
    
    For Each TargetSubfolder In TargetFolder.SubFolders
        
        'フォルダを基準となる列から横にずらして出力することで、サブフォルダの階層を表現
        Call FileSearch(TargetSheet, FolderPath & "\" & TargetSubfolder.Name, ChangePath & "\" & TargetSubfolder.Name, outRow, outColumn + 1, baseColumn, Fso)  '再帰呼出
    
    Next TargetSubfolder
    
    With TargetSheet
        
        For Each TargetFile In TargetFolder.Files

            CurrentFolderPath = FolderPath
            
            '一番下の階層のフォルダから出力して、サブフォルダの階層を表現
            For i = outColumn To baseColumn Step -1
            
                .Cells(outRow, i).Value = Fso.GetBaseName(CurrentFolderPath)
                
                CurrentFolderPath = Fso.GetParentFolderName(CurrentFolderPath)
                    
            Next i
            
            .Cells(outRow, 1) = FolderPath & "\" & TargetFile.Name
            .Cells(outRow, 2) = TargetFile.Name
              
            outRow = outRow + 1
            
        Next TargetFile

    End With
    
    Set TargetFile = Nothing
    Set TargetSubfolder = Nothing
    Set TargetFolder = Nothing
    

End Sub
Public Sub SearchSubFolders_Folder()

Dim Fso As Object
Dim FolderPath As String
Dim ChangePath As String
Dim StartRow As Long
Dim FolderStartColumn As Long
Dim TargetSheet As Worksheet
    
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Set TargetSheet = ThisWorkbook.Sheets("フォルダ名取得2")

    With TargetSheet
    
        .Range("A1").CurrentRegion.Offset(2).ClearContents
        
        FolderPath = .Range("B1").Value
        StartRow = .Range("A3").Row
        FolderStartColumn = .Range("B3").Column
   
    End With
    
    ChangePath = ChangeShortPath(FolderPath)
    
    If ChangePath <> "" Then
    
        Call FolderSearch(TargetSheet, FolderPath, ChangePath, StartRow, FolderStartColumn, FolderStartColumn, Fso)
       
    Else
    
        MsgBox FolderPath & "は存在しません。"
       
    End If
    
    Set Fso = Nothing
    
    
End Sub

Sub FolderSearch(TargetSheet As Worksheet, FolderPath As String, ChangePath As String, outRow As Long, outColumn As Long, baseColumn As Long, Fso As Object)

Dim i As Long
Dim TargetFolder As Object
Dim TargetSubfolder As Object
Dim OriginalPath As String

Dim CurrentFolderPath As String
    
    
    Set TargetFolder = Fso.GetFolder(ChangePath)
    
    For Each TargetSubfolder In TargetFolder.SubFolders
          
        'フォルダを基準となる列から横にずらして出力することで、サブフォルダの階層を表現
        Call FolderSearch(TargetSheet, FolderPath & "\" & TargetSubfolder.Name, ChangePath & "\" & TargetSubfolder.Name, outRow, outColumn + 1, baseColumn, Fso) '再帰呼出
    
    Next TargetSubfolder
        
    With TargetSheet
        
        For Each TargetSubfolder In TargetFolder.SubFolders

            CurrentFolderPath = FolderPath & "\" & TargetSubfolder.Name
            
            '一番下の階層のフォルダから出力して、サブフォルダの階層を表現
            For i = outColumn To baseColumn Step -1
            
                .Cells(outRow, i).Value = Fso.GetBaseName(CurrentFolderPath)
                
                CurrentFolderPath = Fso.GetParentFolderName(CurrentFolderPath)
                    
            Next i
            
            .Cells(outRow, 1) = FolderPath & "\" & TargetSubfolder.Name
            outRow = outRow + 1
            
        Next TargetSubfolder
  
    End With

    Set TargetSubfolder = Nothing
    Set TargetFolder = Nothing
    
End Sub

