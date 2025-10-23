Attribute VB_Name = "SearchSubfolders"
Option Explicit


Public Sub SearchSubFolders_File()

Dim FolderPath As String
Dim StartRow As Long
Dim FolderStartColumn As Long


    With ThisWorkbook.Sheets("ファイル名取得2")
    
        .Range("A1").CurrentRegion.Offset(2).ClearContents
        
        FolderPath = .Range("B1").Value
        
        StartRow = .Range("A3").Row
        FolderStartColumn = .Range("C3").Column
        
    
    End With

    Call FileSearch(FolderPath, StartRow, FolderStartColumn, FolderStartColumn)
    
   
    
End Sub

Sub FileSearch(FolderPath As String, outRow As Long, outColumn As Long, baseColumn As Long)

Dim Fso As Object
Dim f As Object
Dim i As Long

Dim iFolder As Object
Dim OriginalPath As String

Dim CurrentFolderPath As String


    Set Fso = CreateObject("Scripting.FileSystemObject")

    OriginalPath = FolderPath
    
    Set iFolder = Fso.GetFolder(ChangeShortPath(FolderPath))
    
    For Each f In iFolder.SubFolders
    
        Call FileSearch(OriginalPath & "\" & f.Name, outRow, outColumn + 1, baseColumn)  '再帰呼出
    
    Next
        
    With ThisWorkbook.Sheets("ファイル名取得2")
        
        For Each f In iFolder.Files

            CurrentFolderPath = OriginalPath
            
            For i = outColumn To baseColumn Step -1
            
                .Cells(outRow, i).Value = Fso.GetBaseName(CurrentFolderPath)
                
                CurrentFolderPath = Fso.GetParentFolderName(CurrentFolderPath)
                    
            Next i
            
            .Cells(outRow, 1) = OriginalPath & "\" & f.Name
            .Cells(outRow, 2) = f.Name
              
            outRow = outRow + 1
            
        Next f
  
    End With

Set Fso = Nothing

End Sub
Public Sub SearchSubFolders_Folder()

Dim FolderPath As String
Dim StartRow As Long
Dim FolderStartColumn As Long
   
    With ThisWorkbook.Sheets("フォルダ名取得2")
    
        .Range("A1").CurrentRegion.Offset(2).ClearContents
        
        FolderPath = .Range("B1").Value
        StartRow = .Range("A3").Row
        FolderStartColumn = .Range("B3").Column
   
    End With
    
    Call FolderSearch(FolderPath, StartRow, FolderStartColumn, FolderStartColumn)
    
   
    
End Sub

Sub FolderSearch(FolderPath As String, outRow As Long, outColumn As Long, baseColumn As Long)

Dim Fso As Object
Dim f As Object
Dim i As Long

Dim iFolder As Object
Dim OriginalPath As String

Dim CurrentFolderPath As String


    Set Fso = CreateObject("Scripting.FileSystemObject")

    OriginalPath = FolderPath
    
    Set iFolder = Fso.GetFolder(ChangeShortPath(FolderPath))

    For Each f In iFolder.SubFolders
          
        Call FolderSearch(OriginalPath & "\" & f.Name, outRow, outColumn + 1, baseColumn)  '再帰呼出
    
    Next
        
    With ThisWorkbook.Sheets("フォルダ名取得2")
        
        For Each f In iFolder.SubFolders

            CurrentFolderPath = OriginalPath & "\" & f.Name

            For i = outColumn To baseColumn Step -1
            
                .Cells(outRow, i).Value = Fso.GetBaseName(CurrentFolderPath)
                
                CurrentFolderPath = Fso.GetParentFolderName(CurrentFolderPath)
                    
            Next i
            
            .Cells(outRow, 1) = OriginalPath & "\" & f.Name
            outRow = outRow + 1
            
        Next f
  
    End With
    
    Set Fso = Nothing
    
End Sub

