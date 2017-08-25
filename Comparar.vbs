Sub Comparar()
'
' Comparar Macro
' Compara los libros por columnas
    Dim arreglo As Variant
    Dim arreglo2 As Variant
    Dim arreglo3 As Variant
    Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    
    Dim x As String
    Dim y As String
    
    
        Set h1 = Workbooks("book1").Sheets("sheet")
            Set h2 = Workbooks("book2").Sheets("sheet")
    
    For Each aSheet In Worksheets

        Select Case aSheet.Name
                Case "SheetResult"
                Application.DisplayAlerts = False
                aSheet.Delete
                Application.DisplayAlerts = True
    
        End Select
    Next aSheet

    
    arreglo = h1.Range("B2:B97").Value
    arreglo2 = h2.Range("B2:B2046").Value
    contador = 1
    aux = 0
    
    While contador <= UBound(arreglo)
        contador2 = 1
        x = arreglo(contador, 1)
        arreglo3 = Split(x)
  
        Do While contador2 <= UBound(arreglo2)
            y = arreglo2(contador2, 1)
            If InStr(1, y, arreglo3(0)) > 0 Then
                arrList.Add y
            End If
            contador2 = contador2 + 1
        Loop
        contador = contador + 1
    Wend
    ReDim arregloRTA(arrList.Count) As Variant
    
    For Each Item In arrList
        arregloRTA(aux) = Item
        aux = aux + 1
    Next Item
    
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "sheetResult"
    Range("A1") = "NameColumn"
    Range("A2:A4510").Value = Application.Transpose(arregloRTA)
    Columns("A:A").EntireColumn.AutoFit
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
    ordenar
    
    
    
End Sub
Sub borrar()
'
' borrar Macro
' Borrar fila por la fila seleccionada
'
' Acceso directo: CTRL+a
    Selection.EntireRow.Delete
    
  
End Sub
Private Sub ordenar()
   Range("A1").Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes
End Sub

