Attribute VB_Name = "ConvertirLetras"
'declara form tsrpg310 como global
 'Global f As Form
 'Global VGB_Form As Form
 

'   Variables para el proyecto tsrp001
    Global fec_agno              As String
    Global primer_dia_mes        As String

' Variables declaradas para el valor del cheque en palabras
' tsrt13


   Global unidades(10)   As String
   Global decenas(10)   As String
   Global dec_esp(10)   As String
   Global miles(12)   As String
   Global valor_red As String

   Global w_mes As String
   Global aux_dia As String
   Global aux_mes As String
   Global aux_ano As String
   Global mes_pal As String

   Global cantidad(50)  As String

   Global indice     As Integer
   Global primero     As Integer
   Global desde     As Integer
   Global unidad1     As Integer
   Global posicion     As Integer
   Global largo1     As Integer

   Global cant_paso     As String
   Global cantidad_pal     As String
   
 '  Global pos     As Integer
 '   Global total   As Integer
 '  Global cant_pal     As String
                 
 '  Global numero1     As Integer
 '  Global numero2     As Integer
 '  Global numero3     As Integer
 '  Global largo       As String
 '  Global cant1     As String
 '  Global unidad     As String

   Global cantidad_1     As String
   Global cantidad_2     As String

   Global cantidad_valor As String
   Global cant_1     As String
   Global cant_2     As String
   Global l_aux1     As Integer
   Global l_aux2     As Integer
   Global j      As Integer

   Global Valor_chq_en_pal As String

'Otras variables

   Global z_res_sw   As String
   Global respuesta  As Integer
   Global contador  As Integer

' variables para t04

   Global w_presup_caja  As String
   Global w_centro_costo  As String

Global montoescrito As String
Global pesosglobal As String
Global gbl_status As Integer
Global TR_Fld_GloEmpresa As String
Global TR_Fld_GloUndTsr As String

Function centena(gabi As String, pos As Integer, Total As Integer, cant_pal As String, ByVal pPosAct As Integer, ByVal pPosFin As Integer)
   Dim Numero1     As Integer
   Dim numero2     As Integer
   Dim numero3     As Integer
   Dim largo       As String
   Dim cant1     As String
   Dim unidad     As String
   

   cant1 = ""
   unidad = ""

   If pos Mod 2 <> 0 And pos <> 1 And Val(gabi) <> 1 Then
       unidad = miles(pos) + "es"
   Else
       unidad = miles(pos)
   End If

   Numero1 = 0
   numero2 = 0
   numero3 = 0
   largo = Len(gabi)

    If largo = "1" Then
       numero3 = Val(Mid$(gabi, 1, 1))
    End If

    If largo = "2" Then
       numero2 = Val(Mid$(gabi, 1, 1))
       numero3 = Val(Mid$(gabi, 2, 1))
    End If

    If largo = "3" Then
       Numero1 = Val(Mid$(gabi, 1, 1))
       numero2 = Val(Mid$(gabi, 2, 1))
       numero3 = Val(Mid$(gabi, 3, 1))
    End If

  ' 03.03.95 antes If numero1 = 0 And numero2 = 0 And numero3 = 0 Then
   If Numero1 = 0 And numero2 = 0 And numero3 = 0 And pos = 2 Then
      unidad = ""
   End If

 '  If numero1 = 0 Then
      'Return
 '  End If
   Select Case Numero1
       Case 0
       Case 1
           If Numero1 = 1 And (numero2 <> 0 Or numero3 <> 0) Then
               cant1 = "ciento"
           Else
              cant1 = "cien"
           End If
       Case 5
           cant1 = "quinientos"
       Case 7
           cant1 = "setecientos"
       Case 9
           cant1 = "novecientos"
       Case Else
           cant1 = unidades(Numero1) + "cientos"
   End Select

   If numero2 = 1 And numero3 <= 5 Then
      cant1 = cant1 + " " + dec_esp(numero3 + 1)
   Else
       If numero2 <> 0 Then
           cant1 = cant1 + " " + decenas(numero2)
           If numero3 <> 0 Then
             cant1 = cant1 + " y"
           End If
      End If

       If numero3 <> 0 Then
           If Len(cant1) = 0 Then
  '              If numero1 = 0 And numero2 = 0 And numero3 = 1 Then
   '                    cant1 = "UN"
               'Else
               cant1 = unidades(numero3)
    '           End If
           Else
'               If numero1 = 0 And numero2 = 0 And numero3 = 1 Then
'                    cant1 = cant1 + " UN"
'               Else
'                    If numero1 + numero2 > 0 Then
'                        cant1 = cant1 + " UNO"
'                    Else
                        cant1 = cant1 + " " + unidades(numero3)
 '                   End If
'
                    
 '              End If
           
           
               
           End If
       End If
   End If

' Todo lo sgte antes no estaba comentariado
'   If numero1 = 0 And numero2 = 0 And numero3 = 1 And Len(cant_pal) <> 0 Then
'      cant_pal = cant_pal + " " + unidad
'   Else
      cant_pal = cant_pal + " " + cant1 + " " + unidad
'   End If
   
   If pPosAct = pPosFin Then
        If numero2 > 1 And numero3 = 1 Then
            cant_pal = RTrim(cant_pal) + "O"
        End If
        If numero2 = 0 And numero3 = 1 Then
            cant_pal = RTrim(cant_pal) + "O"
        End If
   End If
   
   

   centena = cant_pal
   
End Function

Function format_valor(linea_1 As Integer, linea_2 As Integer, cant As String)

    linea_2 = linea_2 - 7

    l_aux1 = linea_1 + 1

    If Mid$(cant, linea_1, 1) = "" Or Mid$(cant, linea_1, 1) = " " Or Mid$(cant, l_aux1, 1) = "" Or Mid$(cant, l_aux1, 1) = " " Then
        If Mid$(cant, l_aux1, 1) = "" Or Mid$(cant, l_aux1, 1) = " " Then
           l_aux1 = l_aux1 + 1
        End If
        l_aux2 = l_aux1 + linea_2 - 1

        cant_1 = Mid$(cant, 1, linea_1)
        cant_2 = Mid$(cant, l_aux1, l_aux2 - l_aux1 + 1)
    Else
        For I = linea_1 To 1 Step -1
            If Mid$(cant, I, 1) = "" Or Mid$(cant, I, 1) = " " Then

                cant_1 = Mid$(cant, 1, I)

                For j = I + 1 To 300
                    If Mid$(cant, j, 1) <> "" And Mid$(cant, j, 1) <> " " Then
                        l_aux2 = j + linea_2 - 1
                        If l_aux2 > 300 Then
                           l_aux2 = 300
                        End If

                        cant_2 = Mid$(cant, j, l_aux2 - j + 1)
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next I
        j = Len(cant_2)
        I = Len(cant_1)

      'tengo error aqui duplicate definition  l_aux1 = Len(cant)

        If I + j < l_aux1 Then
           I = Len(cant_1)
           j = I + 2
           For I = j To linea_1 - 1
                l_aux2 = I + 3

                If Mid$(cant, I, l_aux2 - I + 1) = "MIL" Then
                   Exit For
                End If

                Let l_aux2 = I + 4

                If Mid$(cant, I, l_aux2 - I + 1) = "DIEZ" Then
                   Exit For
                End If

                aux_can = Mid$(cant, I, 1)
                If aux_can = "A" Or aux_can = "E" Or aux_can = "I" Or aux_can = "O" Or aux_can = "U" Then
                    l_aux2 = I + 1
                    aux_can = Mid$(cant, l_aux2, 1)
                    If aux_can <> "R" And aux_can <> "A" And aux_can <> "E" And aux_can <> "I" And aux_can <> "O" And aux_can <> "U" And aux_can <> "N" Then

                        cant_1 = Mid$(cant, 1, I)

                        l_aux1 = l_aux2
                        l_aux2 = l_aux1 + linea_2 - 1
                        If l_aux2 > 300 Then
                           l_aux2 = 300
                        End If

                       cant_2 = Mid$(cant, l_aux1, l_aux2 - l_aux1 + 1)

                    Else
                       If Mid$(cant, I, l_aux2 - I + 1) = "AR" Then

                           cant_1 = Mid$(cant, 1, I)

                            l_aux1 = l_aux2
                            l_aux2 = l_aux1 + linea_2 - 1
                            If l_aux2 > 300 Then
                               l_aux2 = 300
                            End If

                            cant_2 = Mid$(cant, l_aux1, l_aux2 - l_aux1 + 1)
                        End If
                    End If
                End If

                If Mid$(cant, I, 1) = "N" Then
                    l_aux2 = I + 1
                    If Mid$(cant, l_aux2, 1) = "T" Or Mid$(cant, l_aux2, 1) = "C" Then

                        cant_1 = Mid$(cant, 1, I)
                        l_aux1 = l_aux2
                        l_aux2 = l_aux1 + linea_2 - 1
                        If l_aux2 > 300 Then
                           l_aux2 = 300
                        End If

                        cant_2 = Mid$(cant, l_aux1, l_aux2 - l_aux1 + 1)

                    End If
                End If

                If Mid$(cant, I, 1) = "R" Then
                    l_aux2 = I + 1
                    If Mid$(cant, l_aux2, 1) = "C" Then
                        cant_1 = Mid$(cant, 1, I)
                        l_aux1 = l_aux2
                        l_aux2 = l_aux1 + linea_2 - 1
                        If l_aux2 > 300 Then
                           l_aux2 = 300
                        End If

                       cant_2 = Mid$(cant, l_aux1, l_aux2 - l_aux1 + 1)
                    End If
                End If
            Next I
        End If
    End If

    I = Len(cant_1)
    j = Len(cant_2)

    linea_2 = linea_2 + 7

    If j = 0 Then
        If I < linea_1 - 6 Then
            cant_1 = cant_1 + " 00/100"
            For l_aux2 = I + 8 To linea_1
                cant_1 = cant_1 + "*"
            Next l_aux2
            j = 0
        Else
            cant_2 = "00/100"
            j = 6
        End If
    Else
        j = j + 6
        cant_2 = cant_2 + " 00/100"
    End If

    For I = j + 1 To linea_2
        cant_2 = cant_2 + "*"
    Next I

  ' return   cant_1 , cant_2
End Function

Sub inicializa_var()
    indice = 0
    primero = 0
    desde = 0
    unidad1 = 0
    posicion = 0
    largo1 = 0

    cant_paso = ""
    cantidad_pal = ""

    Numero1 = 0
    numero2 = 0
    numero3 = 0
    largo = ""
    cant1 = ""
    unidad = ""

    cantidad_1 = ""
    cantidad_2 = ""

    cantidad_valor = ""
    cant_1 = ""
    cant_2 = ""
    l_aux1 = 0
    l_aux2 = 0
    I = 0
    j = 0

    Valor_chq_en_pal = ""
End Sub

Sub llena_arreglo()
    unidades(1) = "un"
    unidades(2) = "dos"
    unidades(3) = "tres"
    unidades(4) = "cuatro"
    unidades(5) = "cinco"
    unidades(6) = "seis"
    unidades(7) = "siete"
    unidades(8) = "ocho"
    unidades(9) = "nueve"

    dec_esp(1) = "diez"
    dec_esp(2) = "once"
    dec_esp(3) = "doce"
    dec_esp(4) = "trece"
    dec_esp(5) = "catorce"
    dec_esp(6) = "quince"

    decenas(1) = "diez"
    decenas(2) = "veinte"
    decenas(3) = "treinta"
    decenas(4) = "cuarenta"
    decenas(5) = "cincuenta"
    decenas(6) = "sesenta"
    decenas(7) = "setenta"
    decenas(8) = "ochenta"
    decenas(9) = "noventa"

    miles(1) = ""
    miles(2) = "mil"
    miles(3) = "millon"
    miles(4) = "mil"
    miles(5) = "billon"
    miles(6) = "mil"
    miles(7) = "trillon"
    miles(8) = "mil"
    miles(9) = "cuadrillon"
    miles(10) = "mil"
    miles(11) = "quintillon"
    miles(12) = "mil"
End Sub
Public Function mes_palabras2(w_mes As Integer) As String

  If w_mes = 1 Then
     mes_palabras2 = "ENERO"
  End If

  If w_mes = 2 Then
     mes_palabras2 = "FEBRERO"
  End If

  If w_mes = 3 Then
     mes_palabras2 = "MARZO"
  End If

  If w_mes = 4 Then
     mes_palabras2 = "ABRIL"
  End If

  If w_mes = 5 Then
     mes_palabras2 = "MAYO"
  End If

  If w_mes = 6 Then
     mes_palabras2 = "JUNIO"
  End If

  If w_mes = 7 Then
     mes_palabras2 = "JULIO"
  End If

  If w_mes = 8 Then
     mes_palabras2 = "AGOSTO"
  End If

  If w_mes = 9 Then
     mes_palabras2 = "SEPTIEMBRE"
  End If

  If w_mes = 10 Then
     mes_palabras2 = "OCTUBRE"
  End If
  
  If w_mes = 11 Then
     mes_palabras2 = "NOVIEMBRE"
  End If

  If w_mes = 12 Then
     mes_palabras2 = "DICIEMBRE"
  End If
End Function

Sub mes_palabras(w_mes As String)

  If w_mes$ = "01" Then
     mes_pal$ = "ENERO"
  End If

  If w_mes$ = "02" Then
     mes_pal$ = "FEBRERO"
  End If

  If w_mes$ = "03" Then
     mes_pal$ = "MARZO"
  End If

  If w_mes$ = "04" Then
     mes_pal$ = "ABRIL"
  End If

  If w_mes$ = "05" Then
     mes_pal$ = "MAYO"
  End If

  If w_mes$ = "06" Then
     mes_pal$ = "JUNIO"
  End If

  If w_mes$ = "07" Then
     mes_pal$ = "JULIO"
  End If

  If w_mes$ = "08" Then
     mes_pal$ = "AGOSTO"
  End If

  If w_mes$ = "09" Then
     mes_pal$ = "SEPTIEMBRE"
  End If

  If w_mes$ = "10" Then
     mes_pal$ = "OCTUBRE"
  End If

  If w_mes$ = "11" Then
     mes_pal$ = "NOVIEMBRE"
  End If

  If w_mes$ = "12" Then
     mes_pal$ = "DICIEMBRE"
  End If
End Sub

Public Function monto_palabras(ByVal cantidad As String)
   Dim Deci As String
   Deci = Right(Format(cantidad, "#############0.00"), 2)
   cantidad = CCur(Int(cantidad))
   
   Call llena_arreglo
   cantidad_pal = ""
   largo1 = Len(cantidad)
   desde = 1
   primero = largo1 Mod 3
   unidad1 = Int(largo1 / 3)

   If primero = 0 Then
      primero = 3
   Else
      unidad1 = unidad1 + 1
   End If

   posicion = unidad1
   For indice = primero To largo1 Step 3

       cant_paso = Mid$(cantidad, desde, indice - desde + 1)
       cantidad_pal = centena(cant_paso, posicion, unidad1, cantidad_pal, indice, largo1)
       desde = indice + 1
       posicion = posicion - 1
   Next indice
   monto_palabras = UCase(cantidad_pal) & " CON " & Deci & "/100  "
End Function

Sub VBG_Form_PutMsg(Severity As Integer, Message As String)

  ' tsr003.HelpStatus.Caption = Message
  MsgBox Message, 16
     
End Sub

