Attribute VB_Name = "importaciones"
Public Function AbrirFile(pextension As String, ByRef box As CommonDialog) As String
'IMPLEMENTACION GALLOS
AbrirFile = ""

If Not Cuadro_Dialogo_Abrir(pextension, box) Then
    Exit Function
End If

'Debug.Print box.DefaultExt

If UCase(Right(box.FileName, 3)) <> UCase(Right(pextension, 3)) Then
   MsgBox "La Extensión de archivo no concuerda con el formato elegido", vbCritical, "Archivo Inválido"
   Exit Function
End If
AbrirFile = box.FileName
End Function

Public Function Cuadro_Dialogo_Abrir(pextension As String, ByRef box As CommonDialog) As Boolean
'IMPLEMENTACION GALLOS

 'On Error GoTo ErrHandler
   ' Establece los filtros.
   
   box.CancelError = True
   Select Case pextension
    Case "*.txt"
        box.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|"
    Case "*.dbf"
        box.Filter = "All Files (*.*)|*.*|Tablas Files (*.dbf)|*.dbf|"
    Case "*.mdb"
        box.Filter = "All Files (*.*)|*.*|BD Access (*.mdb)|*.mdb|"
    Case "*.csv"
        box.Filter = "All Files (*.*)|*.*|Microsoft Excel (*.csv)|*.csv|"
    Case "*.xls"
        box.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*"
        '"All Files (*.*)|*.*|Microsoft Excel 97/2000 (*.xls)|*.txt)"
   End Select
   ' Especifique el filtro predeterminado.
   box.FilterIndex = 2
   box.FileName = ""
   box.InitDir = "U:\VPINTO\PLLA2008\" 'App.path
   ' Presenta el cuadro de diálogo Abrir.
   box.ShowOpen
   ' Llamada al procedimiento para abrir archivo.
   Dim pos As String
   
  
   Dim swExiste As Variant
   swExiste = InStr(1, UCase(Trim(box.FileName)), UCase(xFile), vbTextCompare)
   If swExiste = 0 Then
      MsgBox "Archivo Elegido no es el correcto" & Chr(13) & "El Correcto es " & xFile, vbCritical, "Importacion"
      'salir = True
      Txtarchivos = ""
    Else
      Cuadro_Dialogo_Abrir = True
    End If
   Exit Function

ErrHandler:
   Cuadro_Dialogo_Abrir = False
   'El usuario hizo clic en el botón Cancelar.
   Exit Function
End Function
