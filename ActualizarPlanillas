Sub ImportarDatosEntrePlanillas(ByVal sourceFilePath As String, ByVal destFilePath As String, ByVal sourceRange As String, ByVal destRange As String)
    Dim SourceWb As Workbook
    Dim DestWb As Workbook
    
    ' Abre el archivo de origen
    Set SourceWb = Workbooks.Open(sourceFilePath)
    
    ' Abre el archivo de destino
    Set DestWb = Workbooks.Open(destFilePath)
    
    ' Copia los datos del rango de origen
    SourceWb.Sheets(1).Range(sourceRange).Copy
    
    ' Pega los datos en el rango de destino
    DestWb.Sheets(1).Range(destRange).PasteSpecial xlPasteValues
    
    ' Actualiza las fórmulas en el libro de destino
    DestWb.Sheets(1).Calculate
    
    ' Cierra los libros
    SourceWb.Close SaveChanges:=False
    DestWb.Close SaveChanges:=True
End Sub

Sub ImportarDatos()
    Dim sourceFile As String
    Dim destFile As String
    Dim sourceRange As String
    Dim destRange As String
    
    sourceFile = "plaillaA.xls"
    destFile = "planillaB.xls"
    sourceRange = "A3:C100" ' Define el rango de origen
    destRange = "A3" ' Define el rango de destino
    
    ImportarDatosEntrePlanillas sourceFile, destFile, sourceRange, destRange
End Sub

