' Creates a new Excel file from a range of cells.
Sub CreateExcelFileFromRange(rng As Range, filePath As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    rng.Copy wb.Sheets(1).Range("A1")
    wb.SaveAs filePath
    wb.Close
End Sub

' Creates a new Excel file from an array.
Sub CreateExcelFileFromArray(arr As Variant, filePath As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Sheets(1).Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    wb.SaveAs filePath
    wb.Close
End Sub

' Copies files from one directory to another.
Sub CopyFiles(sourcePath As String, destinationPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destinationPath) Then
        fso.CreateFolder destinationPath
    End If
    fso.CopyFile sourcePath & "\*", destinationPath & "\*"
End Sub

' Creates a folder with a name that includes specific date formats.
Sub CreateFolder(path As String, name As String, year As String, month As String, separator As String)
    Dim folderName As String
    folderName = path & "\" & year & separator & month
    If Dir(folderName, vbDirectory) = "" Then
        MkDir folderName
    End If
    MsgBox "Folder created: " & folderName, vbInformation, "Success"
End Sub

' Saves a file with a name that includes specific date formats.
Sub SaveFileWithDate(filePath As String, fileName As String, dateFormat As String)
    Dim fullFilePath As String
    fullFilePath = filePath & "\" & fileName & "_" & Format(Now, dateFormat) & ".xlsx"
    ThisWorkbook.SaveAs fullFilePath
    MsgBox "File saved: " & fullFilePath, vbInformation, "Success"
End Sub
