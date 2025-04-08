Attribute VB_Name = "Funciones_BK"

Public Sub fc_Limpiar_MDB()
On Error GoTo TERMINA
     Dim cat As New ADOX.Catalog
     Dim tb As New ADOX.Table
     Dim cta As New ADODB.Connection
     Dim con As String
     
     con = "provider=microsoft.jet.oledb.4.0;data source=" & _
     nom_BD & ";jet OLEDB:database password="
     cta.Open con

     Set cat.ActiveConnection = cta

     For Each tb In cat.Tables

         If tb.Name <> "MSysACEs" And tb.Name <> "MSysModules" And tb.Name <> "MSysModules2" And tb.Name <> "MSysObjects" And tb.Name <> "MSysQueries" And tb.Name <> "MSysRelationships" Then
            cat.Tables.Delete tb.Name
         End If

         If cat.Tables.count = 6 Then Exit For

     Next

     cta.Close
     Set cta = Nothing
'===========================================
'    Dim tdf As TableDef
'
'
'    Set bd = OpenDatabase(nom_BD, False, False)
'
'    Do While bd.TableDefs.Count > 6
'     For Each tdf In bd.TableDefs
'
'        If tdf.Name <> "MSysACEs" And tdf.Name <> "MSysModules" And tdf.Name <> "MSysModules2" And tdf.Name <> "MSysObjects" And tdf.Name <> "MSysQueries" And tdf.Name <> "MSysRelationships" Then
'          bd.TableDefs.Delete tdf.Name
'        End If
'        If bd.TableDefs.Count = 6 Then Exit For
'     Next tdf
'    Loop
    Exit Sub
    
TERMINA:
    MsgBox "Error : " & Err.Description, vbCritical, "SISTEMA DE PLANILLAS"
'   Call rMensaje("No se pudo abrir la Tabla Temporal en ACCESS, verifique que no esté en uso !!", 1)
 '  frmRDOErrors.Show
   End
End Sub

Public Sub rCarCbo(ByRef pCbo As Control, _
                   ByVal pSql As String, _
                   Optional pTip, _
                   Optional pFmt)
                   
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

    pCbo.Clear

    If (fAbrRst(rs, pSql)) Then
       If (Not rs.EOF) Then
          mLen = Len(rs(0))
          pCbo.Tag = pTip & Format(mLen, "00#") & pFmt
          Do Until rs.EOF
             If pCbo.Name <> "Cmbcta" And pCbo.Name <> "Cmbcta2" Then
                pCbo.AddItem rs(1)
                pCbo.ItemData(pCbo.NewIndex) = rs(0)
             Else
                pCbo.AddItem rs(1) & "  " & rs(0)
             End If
             rs.MoveNext
          Loop
       End If
       If rs.State = 1 Then rs.Close
    End If

End Sub
Public Function fAbrRst(ByRef pRs As ADODB.Recordset, ByVal pSql As String) As Boolean

     Set pRs = cn.Execute(pSql)
     fAbrRst = Not pRs.EOF

End Function

Public Sub rUbiIndCmbBox(ByRef pCbo As Control, ByVal pCod As String, mformat As String)
    ' Ubica ComboBox en posicion coincidente con codigo
    ' Ejm: Call rUbiIndCmbBox(cmbDEPTOS, pCodDpt)
    Dim I As Integer, mLen As Integer, mClv As String, mFmt As String

    ' Caso sin item
    If (pCbo.ListCount = 0) Then
        pCbo.ListIndex = -1
        Exit Sub
    End If
       
    ' Buscar entre los items
    For I = 0 To pCbo.ListCount - 1
        'If (mClv = Left(pCbo.List(i), mLen)) Then
        mClv = Format(pCbo.ItemData(I), mformat)
        If (mClv = pCod) Then
            pCbo.ListIndex = I
            Exit Sub
        End If
    Next

    ' Si no se encontro
    pCbo.ListIndex = -1
End Sub
'*************************************************************************************************
'* Funcion Cargar Combos
'*************************************************************************************************
Sub fc_Descrip_Maestros2(xCodMae As String, xstatus As String, Control As Control)
Dim wciamae As String

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset

sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(xCodMae, 3) & "' and status<>'*' "

If (fAbrRst(rs, sql$)) Then
   If rs!General = "S" Then
      wciamae = " and right(ciamaestro,3)= '" & Right(xCodMae, 3) & "' ORDER BY DESCRIP"
   Else
      wciamae = " and ciamaestro= '" & wcia + Right(xCodMae, 3) & "' ORDER BY DESCRIP"
   End If
End If

If rs.State = 1 Then rs.Close
sql$ = "SELECT cod_maestro2,descrip,flag1 From maestros_2 where status='" & xstatus & "'"
sql$ = sql$ & wciamae

Set rs = cn.Execute(sql$)

Control.Clear
Do While Not rs.EOF
   Control.AddItem Trim(rs!descrip)
   Control.ItemData(Control.NewIndex) = Trim(rs!cod_maestro2)
   rs.MoveNext
Loop

If Right(xCodMae, 3) = "055" And wTipoPla <> "99" Then
   Call rUbiIndCmbBox(Control, wTipoPla, "00")
   If UCase(wuser) <> wAdmin Then Control.Enabled = False
End If
End Sub
Public Function fc_CadDes_Maestros2(xCodMae As String, XCAD As String) As String
Set rs = New ADODB.Recordset
Dim cad As String
cad = "SP_CADDes_Maestros2 '" & xCodMae & "','" & XCAD & "'"
Set rs = cn.Execute(cad)

If Not rs.BOF And Not rs.EOF Then fc_CadDes_Maestros2 = IIf(IsNull(rs!flag1), "", Trim(rs!flag1))
If rs.State = 1 Then rs.Close
End Function
'*************************************************************************************************
'* Funcion Cargar Combos Monedas
'*************************************************************************************************
Sub fc_Descrip_Maestros2_Mon(xCodMae As String, xstatus As String, Control As Control)
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Dim wciamae As String
wciamae = Determina_Maestro(xCodMae)
sql$ = "SELECT cod_maestro2,descrip,flag1 From maestros_2 where status='" & xstatus & "'"
sql$ = sql$ & wciamae
Set rs = cn.Execute(sql$)
Control.Clear
Do While Not rs.EOF
   Control.AddItem Trim(rs!flag1) & "  " & rs!descrip
   Control.ItemData(Control.NewIndex) = Trim(rs!cod_maestro2)
   rs.MoveNext
Loop
End Sub
'*********************************************************************************
'* Funcion Resaltar Texto
'*********************************************************************************
Public Sub ResaltarTexto(texto As Object)
  texto.SelStart = 0
  texto.SelLength = Len(texto)
End Sub
'*********************************************************************************
'* Funcion para validar el ingreso unicamente de digitos del [0-9]
'*********************************************************************************
Public Function fc_ValNumeros(key As Integer) As String
    Select Case key
        Case 0, 8, 13, 27, 32, 45, 48 To 57
            key = key
        Case Else
            key = 0
            MsgBox "Ingrese Dígitos [ 0 - 9 ] únicamente ", vbExclamation, "Verifique"
            Exit Function
        End Select
End Function

'**********************************************************************************************
'* Funcion que devuekve el codigo del dato elegido en un Combobox o ListBox
'**********************************************************************************************

Public Function fc_CodigoComboBox(ByVal NombreControl As Control, ByVal LongCad As Integer) As String
Dim CadCeros As String
If NombreControl.ListIndex <> -1 Then
    Select Case LongCad
        Case 2: CadCeros = "00"
        Case 3: CadCeros = "000"
        Case 4: CadCeros = "0000"
    End Select
    fc_CodigoComboBox = Trim(Right(CadCeros & NombreControl.ItemData(NombreControl.ListIndex), LongCad))
End If
End Function
Public Function lentexto(numero As Integer, texto As String) As String
Dim mLen As Integer
mLen = numero - Len(Trim(texto))
If mLen <> 0 Then
   lentexto = Trim(texto) & Space(mLen)
Else
   lentexto = Trim(texto)
End If
End Function

Public Function fCadNum(pNum As Variant, pFmt As String) As String
   fCadNum = Space(Len(pFmt) - Len(Format(pNum, pFmt))) & Format(pNum, pFmt)
End Function
Public Function SP_PLANILLAS_NOMBRE(xcia As String, xape1 As String, xape2 As String, xnom1 As String, xnom2 As String) As String
Dim xciamae As String
Dim VTipo As String
Dim cod As String

If wTipoPla = "99" Then VTipo = "" Else VTipo = wTipoPla
cod = "01055"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset

sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
sql$ = sql$ & "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(rs, sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If rs.State = 1 Then rs.Close

sql$ = nombre()
sql$ = sql$ & "a.status,a.tipotrabajador,a.fingreso,a.codafp,a.placod,a.codauxinterno,b.descrip " _
    & "from planillas a,maestros_2 b where a.cia='" & wcia & "' and a.status<>'*'"
    sql$ = sql$ & xciamae
    sql$ = sql$ & " and a.tipotrabajador=b.cod_maestro2 and tipotrabajador like '" & Trim(VTipo) + "%" & "'" _
    & "and ap_pat like '" & Trim(xape1) + "%" & "' and ap_mat like '" & Trim(xape2) + "%" & "' and nom_1 like '" & Trim(xnom1) + "%" & "' and nom_1 like '" & Trim(xnom2) + "%" & "'" _
    & "order by nombre"
'    Debug.Print SQL$
    SP_PLANILLAS_NOMBRE = sql$
End Function
Public Function nombre(Optional pAlias As String) As String
Select Case StrConv(gsAdminDB, 1)
       Case Is = "MYSQL"
            sql$ = "SELECT CONCAT(rtrim(ap_pat), ' ',rtrim(ap_mat),' ',rtrim(nom_1),' ',rtrim(nom_2)) AS nombre,"
       Case Is = "SQL SERVER"
            sql$ = "select rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "ap_pat)+' '+rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "ap_mat)+' '+rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "nom_1)+' '+rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "nom_2) as nombre,"
End Select
nombre = sql$
End Function

Public Function SP_PLANILLAS_CODIGO(CODCIA As String, CODPLA As String) As String
Dim xciamae As String
Dim cod As String
Dim VTipo As String
If wTipoPla = "99" Then VTipo = "" Else VTipo = wTipoPla
cod = "01055"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(rs, sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If rs.State = 1 Then rs.Close

sql$ = nombre()
sql$ = sql$ & "a.status,a.tipotrabajador,a.fingreso,a.codafp,a.area,a.placod,a.codauxinterno,b.descrip " _
     & "from planillas a,maestros_2 b where a.status<>'*'"
     sql$ = sql$ & xciamae
     sql$ = sql$ & " and a.tipotrabajador=b.cod_maestro2 and tipotrabajador like '" & Trim(VTipo) + "%" & "'" _
     & "and cia='" & CODCIA & "' AND placod='" & CODPLA & "'" _
     & "order by nombre"
     
SP_PLANILLAS_CODIGO = sql$
End Function
Public Function Fc_Decimals(key As Integer)
Select Case key
       Case 8, 13, 27, 32, 42, 44, 46, 48 To 57
            key = key
       Case Else
            key = 0
            Exit Function
End Select
End Function
Public Function Determina_Maestro(cod) As String
Dim RSdetermina_mae As ADODB.Recordset
Dim xciamae As String
'1055
cn.CursorLocation = adUseClient
Set RSdetermina_mae = New ADODB.Recordset
sql$ = "SELECT GENERAL FROM maestros where " & _
"right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "

'********CENTRO DE COSTOS ?????***************
 If (fAbrRst(RSdetermina_mae, sql$)) Then
   If RSdetermina_mae!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "' ORDER BY cod_maestro2"
      wMaeGen = True
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "' ORDER BY cod_maestro2"
      wMaeGen = False
   End If
End If

If RSdetermina_mae.State = 1 Then RSdetermina_mae.Close
Determina_Maestro = xciamae
End Function
Public Function Determina_Maestro_2(cod) As String
Dim RSdetermina_mae As ADODB.Recordset
Dim xciamae As String
Dim mgeneral As String
cn.CursorLocation = adUseClient
Set RSdetermina_mae = New ADODB.Recordset
sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(RSdetermina_mae, sql$)) Then mgeneral = RSdetermina_mae!General & ""
If RSdetermina_mae.State = 1 Then RSdetermina_mae.Close
If mgeneral = "S" Then
   sql$ = "select general from maestros_2 where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*'"
Else
   sql$ = "select general from maestros_2 where ciamaestro='" & wcia & Right(cod, 3) & "' and status<>'*'"
End If
If (fAbrRst(RSdetermina_mae, sql$)) Then
   If RSdetermina_mae!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "' ORDER BY cod_maestro3"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "' ORDER BY cod_maestro3"
   End If
End If
If RSdetermina_mae.State = 1 Then RSdetermina_mae.Close
Determina_Maestro_2 = xciamae
End Function
Public Function Numero_Hijos(codigo, menor, estudio, fechahoy, edad As Integer) As Integer
Dim RSHijos As ADODB.Recordset
Dim manos As Integer
manos = 0
Numero_Hijos = 0
sql$ = "select * from pladerechohab where cia='" & wcia & "' and placod='" & codigo & "' and escolar='S' and codvinculo='01' and status<>'*'"
If (fAbrRst(RSHijos, sql$)) Then
   Do While Not RSHijos.EOF
      If menor = "S" Then
         manos = perendat(fechahoy, Format(RSHijos!fec_nac, "dd/mm/yyyy"), "a")
         If manos < edad Then Numero_Hijos = Numero_Hijos + 1
      Else
         Numero_Hijos = Numero_Hijos + 1
      End If
      RSHijos.MoveNext
   Loop
   If RSHijos.State = 1 Then RSHijos.Close
End If
End Function
Public Function perendat(fecha1, fecha2, tipo)
' tipo: a=años, m=meses, d=dias
Dim ca As Long, cm As Long, cd As Long
Dim f1 As Variant, f2 As Variant

If CDate(fecha1) < CDate(fecha2) Then
    f1 = fecha1: f2 = fecha2
Else
    f2 = fecha1: f1 = fecha2
End If
ca = DateDiff("yyyy", f1, f2)
If Format(f2, "mmdd") < Format(f1, "mmdd") Then
    ca = ca - 1
End If
cm = DateDiff("m", f1, f2) - (ca * 12)
cd = DateDiff("d", Format(f1, "dd"), Format(f2, "dd"))
If cd < 0 Then
    cm = cm - 1
cd = DateDiff("d", DateSerial(Year(f2), Month(f2) - 1, Day(f1)), f2)
End If
Select Case tipo
Case "d"
    perendat = cd
Case "m"
    perendat = cm
Case "a"
    perendat = ca
End Select
End Function

Public Function Name_Month(mes As String) As String
Select Case mes
       Case Is = "01"
            Name_Month = "ENERO"
       Case Is = "02"
            Name_Month = "FEBRERO"
       Case Is = "03"
            Name_Month = "MARZO"
       Case Is = "04"
            Name_Month = "ABRIL"
       Case Is = "05"
            Name_Month = "MAYO"
       Case Is = "06"
            Name_Month = "JUNIO"
       Case Is = "07"
            Name_Month = "JULIO"
       Case Is = "08"
            Name_Month = "AGOSTO"
       Case Is = "09"
            Name_Month = "SETIEMBRE"
       Case Is = "10"
            Name_Month = "OCTUBRE"
       Case Is = "11"
            Name_Month = "NOVIEMBRE"
       Case Is = "12"
            Name_Month = "DICIEMBRE"
End Select
End Function
Public Function DestinoPort(ByVal wSalida As String, ByVal wFileSal As String, mshell As Boolean)
Dim X As Printer
For Each X In Printers
    If X.DeviceName = wSalida Then
       Set Printer = X
    Exit For
    End If
Next

'*********CODIGO PARA RODA*************************************************
If mshell = True Then
   Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & X.DeviceName, 0)
Else
   CopyFile wFileSal, Printer.Port, FILE_NOTIFY_CHANGE_LAST_WRITE
End If
'**************************************************************************

'*********CODIGO PARA PPM**************************************************
'If Left(Trim(X.DeviceName), 2) = "EP" Then
 '   Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & "LPT1", 0)
'Else
 '  Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & "\\roberto\epsonfx-", 0)
'End If
'**************************************************************************

End Function

Public Function Ultimo_Dia(mes As Integer, ano As Integer) As Integer
Select Case mes
       Case Is = 1, 3, 5, 7, 8, 10, 12
            Ultimo_Dia = 31
       Case Is = 2
            If (ano Mod 4) = 0 Then Ultimo_Dia = 29 Else Ultimo_Dia = 28
       Case Is = 4, 6, 9, 11
            Ultimo_Dia = 30
End Select
End Function
Public Function Imprime_Txt(March As String, ver As String)
wPrintFile = March
Formimpri.Lblver = ver
Load Formimpri
Formimpri.Show
Formimpri.ZOrder 0

End Function
Public Function Remun_Horas(codigo As String) As Integer
Select Case codigo
       Case Is = "01"
            Remun_Horas = 1
       Case Is = "08"
            Remun_Horas = 13
       Case Is = "09"
            Remun_Horas = 2
       Case Is = "10"
            Remun_Horas = 10
       Case Is = "11"
            Remun_Horas = 11
       Case Is = "21"
            Remun_Horas = 17
       Case Is = "22"
            Remun_Horas = 19
       Case Is = "24"
            Remun_Horas = 20
       Case Is = "25"
            Remun_Horas = 21
       Case Is = "26"
            Remun_Horas = 22
       Case Is = "27"
            Remun_Horas = 23
       Case Is = "28"
            Remun_Horas = 24
       Case Is = "29"
            Remun_Horas = 25
       Case Is = "30"
            Remun_Horas = 7
       'Case Is = "17"
        '    Remun_Horas = 17
       Case Else
            Remun_Horas = 0
End Select

End Function
Public Function No_Apostrofe(Tecla As Integer)
If Tecla = 39 Then No_Apostrofe = 0 Else No_Apostrofe = Tecla
End Function
'*********************************************************************************
'* Funcion para validar el ingreso unicamente de digitos numericos mas el punto decimal
'*********************************************************************************
Public Function fc_ValDecimal(key As Integer) As String
    Select Case key
        Case 0, 8, 13, 27, 32, 45, 46, 48 To 57
            key = key
        Case Else
            key = 0
            MsgBox "Ingrese Dígitos [ 0 - 9 ] únicamente ", vbExclamation, "Verifique"
            Exit Function
        End Select
End Function
Public Function fMaxDay(ByVal pMes As Integer, ByVal pAno As Integer) As String
   Dim cad As String
   
   If pAno Mod 4 = 0 Then
    cad = "312931303130313130313031"
   Else
    cad = "312831303130313130313031"
   End If
   
   If pMes = 0 Then pMes = 12
   fMaxDay = Mid(cad, 2 * pMes - 1, 2)
   
End Function
Public Function Compara_Fechas(f1 As String, f2 As String) As Boolean
'Año
If Val(Mid(f1, 7, 4)) > Val(Mid(f2, 7, 4)) Then
   Compara_Fechas = True
   Exit Function
End If
If Val(Mid(f1, 7, 4)) < Val(Mid(f2, 7, 4)) Then
   Compara_Fechas = False
   Exit Function
End If
'Mes
If Val(Mid(f1, 4, 2)) > Val(Mid(f2, 4, 2)) Then
   Compara_Fechas = True
   Exit Function
End If
If Val(Mid(f1, 4, 2)) < Val(Mid(f2, 4, 2)) Then
   Compara_Fechas = False
   Exit Function
End If
'Dia
If Val(Mid(f1, 1, 2)) > Val(Mid(f2, 1, 2)) Then
   Compara_Fechas = True
   Exit Function
End If
If Val(Mid(f1, 1, 2)) < Val(Mid(f2, 1, 2)) Then
   Compara_Fechas = False
   Exit Function
End If
If Val(Mid(f1, 1, 2)) = Val(Mid(f2, 1, 2)) Then
   Compara_Fechas = False
   Exit Function
End If
End Function
Public Function lentextosp(numero As Integer, texto As String) As String
Dim mLen As Integer
mLen = numero - Len(texto)
If mLen <> 0 Then
   lentextosp = Trim(texto) & Space(mLen)
Else
   lentextosp = texto
End If
End Function

Public Function Tipo_Planilla() As String
Dim RStp As ADODB.Recordset
Set RStp = New ADODB.Recordset
sql$ = "select * from users where sistema='04' and status<>'*' and cod_cia='" & wcia & "' and name_user='" & wuser & "'"
If (fAbrRst(RStp, sql)) Then
   Tipo_Planilla = RStp!TIPOPLA & ""
   If RStp!admin = "1" Then Tipo_Planilla = "99"
Else
   Tipo_Planilla = ""
End If
RStp.Close
End Function
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
Function centena(gabi As String, pos As Integer, Total As Integer, cant_pal As String, ByVal pPosAct As Integer, ByVal pPosFin As Integer)
   Dim numero1     As Integer
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

   numero1 = 0
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
       numero1 = Val(Mid$(gabi, 1, 1))
       numero2 = Val(Mid$(gabi, 2, 1))
       numero3 = Val(Mid$(gabi, 3, 1))
    End If

  ' 03.03.95 antes If numero1 = 0 And numero2 = 0 And numero3 = 0 Then
   If numero1 = 0 And numero2 = 0 And numero3 = 0 And pos = 2 Then
      unidad = ""
   End If

 '  If numero1 = 0 Then
      'Return
 '  End If
   Select Case numero1
       Case 0
       Case 1
           If numero1 = 1 And (numero2 <> 0 Or numero3 <> 0) Then
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
           cant1 = unidades(numero1) + "cientos"
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
   cant_pal = cant_pal + " " + cant1 + " " + unidad
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
Public Function AsteriscoR(numero As Integer, texto As String) As String
Dim mLen As Integer
mLen = numero - Len(Trim(texto))
If mLen <> 0 Then
  AsteriscoR = Trim(texto) & String(mLen, "*")
   
Else
   AsteriscoR = Trim(texto)
End If
End Function
Public Function Fecha_Promedios(mFactProm As Integer, fecproc As String) As String
      If mFactProm > 12 Then
         masprom = Int(mFactProm / 12)
         X = 0
         X = mFactProm - Val(Mid(fecproc, 4, 2))
         If X < 0 Then
            fprom = "01/"
            fprom = fprom & Format((X + (12 * masprom)), "00") & "/"
            fprom = fprom & Format(Val(Mid(fecproc, 7, 4)) - masprom, "0000")
         Else
         End If
      ElseIf mFactProm = 12 Then
         fprom = "01/"
         fprom = fprom & Mid(fecproc, 4, 2) & "/"
         fprom = fprom & Format(Val(Mid(fecproc, 7, 4)) - 1, "0000")
      Else
         If Val(Mid(fecproc, 4, 2)) > mFactProm Then
            fprom = "01/" & Format(Val(Mid(fecproc, 4, 2)) - mFactProm, "00") & "/" & Mid(fecproc, 7, 4)
         Else
            fprom = "01/"
            fprom = fprom & Format((12 - (mFactProm - Val(Mid(fecproc, 4, 2)))), "00") & "/"
            fprom = fprom & Format(Val(Mid(fecproc, 7, 4)) - 1, "0000")
         End If
      End If
Fecha_Promedios = fprom
End Function


Public Sub CargaCombo(ByRef pCombo As Object, Optional pCodMaestro As String, Optional pSql As String, Optional pTodos As Boolean = False, Optional pNinguno As Boolean = False, Optional pAutoCorrela As Boolean = False)
Dim sSQL As String
Dim rs As Object
Dim correla As String

If Len(Trim(pCodMaestro)) > 0 Then
    sSQL = " select cod_maestro2,descrip from maestros_2 where ciamaestro='" & Trim(CODCIA) & (pCodMaestro) & "'"
Else
    sSQL = pSql
End If

If Trim(sSQL) <> "" Then
Set rs = cn.Execute(sSQL)

    If rs.Fields.count > 0 Then
        pCombo.Clear
        Do While Not rs.EOF
            correla = Right(rs(0), 4)
            
            If pAutoCorrela Then
                pCombo.AddItem Trim(rs(1)), , Val(correla)
            Else
                pCombo.AddItem Trim(rs(1)), , Val(rs(0))
            End If
            
            rs.MoveNext
        Loop
            
        rs.Close
    End If
End If

If pNinguno Then
    pCombo.AddItem Trim("<Ninguno>"), , 0
End If

If pTodos Then
    pCombo.AddItem Trim("<Todos>"), , 999
End If

Set rs = Nothing

End Sub

Public Function Semana_Calcular(ByVal pSemana As Integer, ByVal pSemanaPago As Integer, _
AñoProceso As String, CodEmpresa As String) As Boolean

'************codigo nuevo agregado giovanni 13092007***************************
Dim rs_Semana_Calculo As ADODB.Recordset
Dim s_Semana_Calculo As String

'********************************************************************************
' MODIFICADO POR RICARDO HINOSTROZA
' FECHA DE MODIFICACION : 07/01/2008
' SE MODIFICA PSEMANA PARA OBTENER DATOS
'********************************************************************************

Dim PCADENA As String
'If CodEmpresa = "05" Then
    If pSemana < 10 Then
        PCADENA = "' and semana='0" & pSemana & "' "
       Else
       PCADENA = "' and semana='" & pSemana & "' "
    End If
'Else
'    PCADENA = "' and semana='" & pSemana & "' "
'End If

s_Semana_Calculo = "select Calculo_Asigfam from plasemanas where cia='" & CodEmpresa & "' " & _
"and status <> '*' and ano='" & AñoProceso & PCADENA

Set rs_Semana_Calculo = New ADODB.Recordset
rs_Semana_Calculo.Open s_Semana_Calculo, cn, adOpenKeyset, adLockOptimistic

If Not rs_Semana_Calculo.EOF Then
    If rs_Semana_Calculo!Calculo_Asigfam = "S" Then
        Semana_Calcular = True
    Else
        Semana_Calcular = False
    End If
Else
    Semana_Calcular = False
End If
Set rs_Semana_Calculo = Nothing

''Dim resultado As Integer
'''****************codigo agregado giovanni 29082007************************
''Dim Psemana2 As Integer
''Psemana2 = pSemana
'''*************************************************************************
''
''Select Case pSemanaPago
''Case 1
''    If pSemanaPago < pSemana Then
''        resultado = pSemana Mod 4
''        If resultado = pSemana Then Semana_Calcular = False: Exit Function
''        If resultado = 0 Then Semana_Calcular = True: Exit Function
''        If resultado <> pSemanaPago + 1 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    Else
''        If pSemanaPago - pSemana <> 0 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    End If
''Case 2
''    If pSemanaPago < pSemana Then
''        '************codigo modificado giovanni 29082007******************************
''        resultado = pSemana Mod 4
''        If resultado = pSemana Then Semana_Calcular = False: Exit Function
''        If resultado = 0 Then Semana_Calcular = True: Exit Function
''Procesa_Semana:
''        If Psemana2 > 5 Then
''        Psemana2 = Psemana2 - 5
''        GoTo Procesa_Semana
''        End If
''        resultado = Psemana2
''        If resultado <> pSemanaPago Then Semana_Calcular = False: Exit Function
''       If resultado = pSemanaPago Then Semana_Calcular = True: Exit Function
''    '****************codigo ingresado PPM giovanni 29082007***********************************
''    'select Case pSemana
''     '   Case Is = 2, 7, 11, 15, 20, 24, 29, 33, 37, 41, 46, 50
''      '      Semana_Calcular = True
''       ' Case Else
''        '    Semana_Calcular = False
''    'End Select
''    '*************************************************************************************
''    Else
''        If pSemanaPago - pSemana <> 0 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    End If
''Case 3
''    If pSemanaPago < pSemana Then
''        resultado = pSemana Mod 4
''        If resultado = pSemana Then Semana_Calcular = False: Exit Function
''        If resultado = 0 Then Semana_Calcular = True: Exit Function
''        If resultado <> pSemanaPago + 1 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    Else
''        If pSemanaPago - pSemana <> 0 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    End If
''Case 4
''    If pSemanaPago < pSemana Then
''        resultado = pSemana Mod 4
''        If resultado = pSemana Then Semana_Calcular = False: Exit Function
''        If resultado = 0 Then Semana_Calcular = True: Exit Function
''        If resultado <> pSemanaPago + 1 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    Else
''        If pSemanaPago - pSemana <> 0 Then
''            Semana_Calcular = False
''        Else
''            Semana_Calcular = True
''        End If
''    End If
''End Select

End Function
