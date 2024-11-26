Attribute VB_Name = "M�dulo1"
Sub InsertarImagenes()
    ' Seleccionar la hoja
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hoja 1") ' Cambia "Hoja 1" al nombre de tu hoja si es necesario

    ' Configuraci�n: Ajusta estas variables seg�n tus necesidades
    Dim startCellAddress As String
    startCellAddress = "E3"  ' Direcci�n de la celda inicial para insertar im�genes
    
    Dim startNumber As Integer
    startNumber = 301  ' N�mero inicial de la imagen
    
    Dim endNumber As Integer
    endNumber = 412  ' N�mero final de la imagen

    ' Ruta base de las im�genes
    Dim basePath As String
    basePath = "C:\Users\Falcon\Documents\Falcon\Scripts\Etiquetas\img\qr_"

    ' Establecer la celda inicial a partir de la direcci�n proporcionada
    Dim startCell As Range
    Set startCell = ws.Range(startCellAddress)
    
    ' Bucle para insertar im�genes desde el n�mero inicial hasta el n�mero final
    Dim i As Integer
    For i = startNumber To endNumber
        ' Construir la ruta de la imagen actual
        Dim imagePath As String
        imagePath = basePath & i & ".png"
        
        ' Seleccionar la celda de destino
        startCell.Offset(i - startNumber, 0).Select
        
        ' Insertar la imagen en la celda seleccionada
        Selection.InsertPictureInCell (imagePath)
    Next i
End Sub

