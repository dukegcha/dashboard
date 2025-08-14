Attribute VB_Name = "Module11"
Sub First()
    ' Remove the first 8 rows
    Rows("1:8").Delete Shift:=xlUp
    

    

End Sub

Sub Second()
        ' Remove columns A and C
    Columns("A").Delete Shift:=xlToLeft
    
    

End Sub
Sub SecondR()
        ' Remove columns A and C
    Columns("A").Delete Shift:=xlToLeft
    Columns("A").Delete Shift:=xlToLeft
    
    

End Sub

Sub Third()
    Columns("B").Delete Shift:=xlToLeft
End Sub


Sub Fourth()
    Rows("2").Delete Shift:=xlUp
End Sub

Sub ReorderColumnsByName()
    Dim col As Range
    Dim colOrder As Variant
    Dim i As Integer
    Dim colPos As Integer
    Dim targetPos As Integer
    
    ' Define the desired order of columns by their names
    colOrder = Array("Material", "Delivery #", "ShpPoint", "Type", "Ac.GI date", "Quantity", "         Volume", "Division", "[WE]State")
    
    ' Loop through each column in the desired order
    For i = LBound(colOrder) To UBound(colOrder)
        ' Find the column with the header name from the list
        For Each col In Rows(1).Cells
            If col.Value = colOrder(i) Then
                colPos = col.Column ' Get the column number
                Exit For
            End If
        Next col
        
        ' Only move columns that are not already in the desired position
        If colPos > 0 Then
            targetPos = i + 1  ' Position where the column needs to go
            
            ' Only cut and insert if the current position is not the target
            If colPos <> targetPos Then
                Columns(colPos).Cut
                Columns(targetPos).Insert Shift:=xlToRight
            End If
        End If
    Next i
End Sub

Sub RemoveBlanks()
    Dim lastRow As Long
    Dim i As Long
    Dim typeColumn As Integer
    
    ' Find the "Type" column by its header in row 1
    typeColumn = 0
    For i = 1 To Columns.Count
        If Cells(1, i).Value = "Type" Then
            typeColumn = i
            Exit For
        End If
    Next i
    
    ' If "Type" column is found
    If typeColumn > 0 Then
        ' Find the last row with data in the "Type" column
        lastRow = Cells(Rows.Count, typeColumn).End(xlUp).Row
        
        ' Loop through each cell in the "Type" column from bottom to top
        For i = lastRow To 2 Step -1 ' Start from row 2 to skip header row
            If IsEmpty(Cells(i, typeColumn)) Then
                Rows(i).Delete Shift:=xlUp
            End If
        Next i
    Else
        MsgBox """Type"" column not found.", vbExclamation
    End If
End Sub

Sub SaveAsXlsx()
    Dim filePath As String
    Dim fileName As String
    Dim baseName As String
    Dim todayDate As String
    
    ' Define the folder where you want to save the CSV
    filePath = "C:\Users\duke.cha\Desktop\all_gi\cleaned\"
    
    ' Ensure the folder exists
    If Dir(filePath, vbDirectory) = "" Then
        MsgBox "The folder does not exist: " & filePath, vbExclamation
        Exit Sub
    End If
    
    ' Get today's date in YYYYMMDD format
    todayDate = Format(Date, "YYYYMMDD")
    
    ' Get the base name of the workbook (remove extension)
    baseName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    
    fileName = baseName & "_" & todayDate & ".xlsx"
    
    
    ' Save the active sheet as a CSV file in the specified folder
    ActiveSheet.SaveAs fileName:=filePath & fileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    
End Sub
Sub SaveAscsv_return()
    Dim filePath As String
    Dim fileName As String
    Dim baseName As String
    Dim todayDate As String
    
    ' Define the folder where you want to save the CSV
    filePath = "C:\Users\duke.cha\Desktop\return\cleaned\"
    
    ' Ensure the folder exists
    If Dir(filePath, vbDirectory) = "" Then
        MsgBox "The folder does not exist: " & filePath, vbExclamation
        Exit Sub
    End If
    
    ' Get today's date in YYYYMMDD format
    todayDate = Format(Date, "YYYYMMDD")
    
    ' Get the base name of the workbook (remove extension)
    baseName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    
    fileName = baseName & "_" & todayDate & ".csv"
    
    
    ' Save the active sheet as a CSV file in the specified folder
   ActiveSheet.SaveAs fileName:=filePath & fileName, FileFormat:=xlCSVUTF8, CreateBackup:=False

    
    
End Sub
Sub SaveAscsv_gr()
    Dim filePath As String
    Dim fileName As String
    Dim baseName As String
    Dim todayDate As String
    
    ' Define the folder where you want to save the CSV
    filePath = "C:\Users\duke.cha\Desktop\GR\cleaned\"
    
    ' Ensure the folder exists
    If Dir(filePath, vbDirectory) = "" Then
        MsgBox "The folder does not exist: " & filePath, vbExclamation
        Exit Sub
    End If
    
    ' Get today's date in YYYYMMDD format
    todayDate = Format(Date, "YYYYMMDD")
    
    ' Get the base name of the workbook (remove extension)
    baseName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    
    fileName = baseName & "_" & todayDate & ".csv"
    
    
    ' Save the active sheet as a CSV file in the specified folder
   ActiveSheet.SaveAs fileName:=filePath & fileName, FileFormat:=xlCSVUTF8, CreateBackup:=False

    
    
End Sub

Sub RunAll()
    Call First
    Call Second
    Call Third
    Call Fourth
    Call ReorderColumnsByName
    Call RemoveBlanks
    Call SaveAsXlsx
End Sub


Sub Runlogtrend()
    Call First
    Call Second
    Call Third
    Call Fourth
End Sub

Sub savelograw()
    Call First
    Call Second
    Call Third
    Call Fourth
    Call SaveAsXlsx
End Sub

Sub return_clean()
    Call First
    Call SecondR
    Call Fourth
End Sub


