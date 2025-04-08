Attribute VB_Name = "Funciones"
' Estados del ADODB
Global Const StateClosed = 0
Global Const StateOpen = 1
Global Const StateFetching = 8
Global Const StateExecuting = 4
Global Const StateConnecting = 2

' Tipo de Cursores
Global Const TypeDynamic = 2
Global Const TypeForwardOnly = 0
Global Const TypeKeyset = 1
Global Const TypeStatic = 3

' Ubicacion de Cursor
Global Const LocationServer = 2
Global Const LocationCliente = 3

' Razon Social CIA
Global Const cNomCiaLargo = 0
Global Const cNomCiaCorto = 1

Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim xlSheet As Object
Dim xlApp2 As Object

'Dim xlSheet As Excel.Worksheet
'Dim xlApp2 As Excel.Application

Dim rsdepo As New Recordset
Dim rs2 As ADODB.Recordset

Public Sub fc_Limpiar_MDB()
On Error GoTo Termina
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
    
Termina:
    MsgBox "Error : " & Err.Description, vbCritical, "SISTEMA DE PLANILLAS"
'   Call rMensaje("No se pudo abrir la Tabla Temporal en ACCESS, verifique que no esté en uso !!", 1)
 '  frmRDOErrors.Show
   End
End Sub

Public Sub rCarCbo(ByRef pCbo As Control, _
                   ByVal pSql As String, _
                   Optional pTip, _
                   Optional pFmt)
                   
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

    pCbo.Clear

    If (fAbrRst(Rs, pSql)) Then
       If (Not Rs.EOF) Then
          mLen = Len(Rs(0))
          pCbo.Tag = pTip & Format(mLen, "00#") & pFmt
          Do Until Rs.EOF
             If pCbo.Name <> "Cmbcta" And pCbo.Name <> "Cmbcta2" Then
                pCbo.AddItem Rs(1)
                pCbo.ItemData(pCbo.NewIndex) = Rs(0)
             Else
                pCbo.AddItem Rs(1) & "  " & Rs(0)
             End If
             Rs.MoveNext
          Loop
       End If
       If Rs.State = 1 Then Rs.Close
    End If

End Sub
Public Sub rCarListBox(ByRef pListBox As Control, _
                   ByVal pSql As String, _
                   Optional pTip, _
                   Optional pFmt)
                   
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

    pListBox.Clear

    If (fAbrRst(Rs, pSql)) Then
       If (Not Rs.EOF) Then
          Do Until Rs.EOF
                pListBox.AddItem Rs(1)
                pListBox.ItemData(pListBox.NewIndex) = Rs(0)
             Rs.MoveNext
          Loop
       End If
       
       
       If Rs.State = 1 Then Rs.Close
    End If

End Sub

Public Function fAbrRst(ByRef pRs As ADODB.Recordset, ByVal pSql As String) As Boolean
     cn.CommandTimeout = 0
     
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
Sub fc_Descrip_Maestros2(xCodMae As String, xstatus As String, Control As Control, Optional pSoloCodSunat As Boolean = False, Optional pSoloBancoHaber As Boolean = False)
Dim wciamae As String
Dim sValor As String

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset

Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(xCodMae, 3) & "' and status<>'*' "

If (fAbrRst(Rs, Sql$)) Then
   If Rs!General = "S" Then
      wciamae = " and right(ciamaestro,3)= '" & Right(xCodMae, 3) & "'"
   Else
      wciamae = " and ciamaestro= '" & wcia + Right(xCodMae, 3) & "'"
   End If
End If

If Rs.State = 1 Then Rs.Close
If xCodMae = "01070" Then
    Sql$ = "SELECT cod_maestro2,Case cod_maestro2 when '01' then 'ZORRITOS' else descrip end as descrip,flag1 From maestros_2 where status='" & xstatus & "'"
Else
   Sql$ = "SELECT cod_maestro2,descrip,flag1 From maestros_2 where status='" & xstatus & "'"
End If
If pSoloCodSunat Then
    Sql$ = Sql$ & " and isnull(codsunat,'')<>''"
End If
If Right(xCodMae, 3) = "007" Then
    Sql$ = Sql$ & " and flag8='1'"
    If pSoloBancoHaber = True Then
        Sql$ = Sql$ & " and cod_maestro2 in(select banco from PlaBcoCta where cia='" & wcia & "' and status !='*')"
    End If
End If

Sql$ = Sql$ & wciamae

If xCodMae = "01070" Then
   Sql$ = Sql$ & " Union all Select '00','CANTERA',''" & " ORDER BY DESCRIP"
Else
    Sql$ = Sql$ & " ORDER BY DESCRIP"
End If

Set Rs = cn.Execute(Sql$)

Control.Clear
Do While Not Rs.EOF
   If Right(xCodMae, 3) = "006" Then
        Control.AddItem Mid(Trim(Rs!flag1), 2, 3)
   Else
        Control.AddItem Trim(Rs!DESCRIP)
        Control.ItemData(Control.NewIndex) = Trim(Rs!cod_maestro2)
   End If
   Rs.MoveNext
Loop

If Right(xCodMae, 3) = "055" And wTipoPla <> "99" Then
   Call rUbiIndCmbBox(Control, wTipoPla, "00")
   If UCase(wuser) <> wAdmin Then Control.Enabled = False
End If
End Sub
Public Function fc_CadDes_Maestros2(xCodMae As String, XCAD As String) As String
Set Rs = New ADODB.Recordset
Dim cad As String
cad = "SP_CADDes_Maestros2 '" & xCodMae & "','" & XCAD & "'"
Set Rs = cn.Execute(cad)

If Not Rs.BOF And Not Rs.EOF Then fc_CadDes_Maestros2 = IIf(IsNull(Rs!flag1), "", Trim(Rs!flag1))
If Rs.State = 1 Then Rs.Close
End Function
'*************************************************************************************************
'* Funcion Cargar Combos Monedas
'*************************************************************************************************
Sub fc_Descrip_Maestros2_Mon(xCodMae As String, xstatus As String, Control As Control)
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Dim wciamae As String
wciamae = Determina_Maestro(xCodMae)
Sql$ = "SELECT cod_maestro2,descrip,flag1 From maestros_2 where status='" & xstatus & "'"
Sql$ = Sql$ & wciamae
Set Rs = cn.Execute(Sql$)
Control.Clear
Do While Not Rs.EOF
   Control.AddItem Trim(Rs!flag1) & "  " & Rs!DESCRIP
   Control.ItemData(Control.NewIndex) = Trim(Rs!cod_maestro2)
   Rs.MoveNext
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
Public Function fc_ValNumeros(Key As Integer) As String
    Select Case Key
        Case 0, 8, 13, 27, 32, 45, 48 To 57
            Key = Key
        Case Else
            Key = 0
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
Public Function fc_CodigoIntListBox(ByVal NombreControl As Control) As Integer
If NombreControl.ListIndex <> -1 Then
    fc_CodigoIntListBox = NombreControl.ItemData(NombreControl.ListIndex)
End If
End Function

Public Function lentexto(numero As Integer, texto As String) As String
Dim mLen As Integer
mLen = numero - Len(Mid(Trim(texto), 1, numero))
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
Set Rs = New ADODB.Recordset

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(Rs, Sql$)) Then
   If Rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If Rs.State = 1 Then Rs.Close

Sql$ = nombre()
Sql$ = Sql$ & "a.status,a.tipotrabajador,a.fingreso,a.codafp,a.placod,a.codauxinterno,b.descrip " _
    & "from planillas a,maestros_2 b where a.cia='" & wcia & "' and a.status<>'*'"
    Sql$ = Sql$ & xciamae
    Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 and tipotrabajador like '" & Trim(VTipo) + "%" & "'" _
    & "and ap_pat like '" & Trim(xape1) + "%" & "' and ap_mat like '" & Trim(xape2) + "%" & "' and nom_1 like '" & Trim(xnom1) + "%" & "' and nom_1 like '" & Trim(xnom2) + "%" & "'" _
    & "order by nombre"
'    Debug.Print SQL$
    SP_PLANILLAS_NOMBRE = Sql$
End Function
Public Function nombre(Optional pAlias As String) As String
Select Case StrConv(gsAdminDB, 1)
       Case Is = "MYSQL"
            Sql$ = "SELECT CONCAT(rtrim(ap_pat), ' ',rtrim(ap_mat),' ',rtrim(nom_1),' ',rtrim(nom_2)) AS nombre,"
       Case Is = "SQL SERVER"
            Sql$ = "select rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "ap_pat)+' '+rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "ap_mat)+' '+rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "nom_1)+' '+rtrim(" & IIf(Len(Trim(pAlias)) > 0, pAlias & ".", "") & "nom_2) as nombre,"
End Select
nombre = Sql$
End Function

Public Function SP_PLANILLAS_CODIGO(CODCIA As String, CODPLA As String) As String
Dim xciamae As String
Dim cod As String
Dim VTipo As String
If wTipoPla = "99" Then VTipo = "" Else VTipo = wTipoPla
cod = "01055"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(Rs, Sql$)) Then
   If Rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If Rs.State = 1 Then Rs.Close

Sql$ = nombre()
Sql$ = Sql$ & "a.status,a.tipotrabajador,a.fingreso,a.codafp,a.area,a.placod,a.codauxinterno,b.descrip " _
     & "from planillas a,maestros_2 b where a.status<>'*'"
     Sql$ = Sql$ & xciamae
     Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 and tipotrabajador like '" & Trim(VTipo) + "%" & "'" _
     & "and cia='" & CODCIA & "' AND placod like '" & CODPLA + "%" & "'" _
     & "order by nombre"
     
SP_PLANILLAS_CODIGO = Sql$
End Function
Public Function Fc_Decimals(Key As Integer)
Select Case Key
       Case 8, 13, 27, 32, 42, 44, 46, 48 To 57
            Key = Key
       Case Else
            Key = 0
            Exit Function
End Select
End Function
Public Function Determina_Maestro(cod) As String
Dim RSdetermina_mae As ADODB.Recordset
Dim xciamae As String
'1055
cn.CursorLocation = adUseClient
Set RSdetermina_mae = New ADODB.Recordset
Sql$ = "SELECT GENERAL FROM maestros where " & _
"right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "

'********CENTRO DE COSTOS ?????***************
 If (fAbrRst(RSdetermina_mae, Sql$)) Then
   If RSdetermina_mae!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "' "
      wMaeGen = True
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "' "
      wMaeGen = False
   End If
   If Mid(cod, 3, 3) = "077" Then
      xciamae = xciamae & " order by convert(integer,flag7)"
   Else
      xciamae = xciamae & " ORDER BY cod_maestro2"
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
Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(RSdetermina_mae, Sql$)) Then mgeneral = RSdetermina_mae!General & ""
If RSdetermina_mae.State = 1 Then RSdetermina_mae.Close
If mgeneral = "S" Then
   Sql$ = "select general from maestros_2 where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*'"
Else
   Sql$ = "select general from maestros_2 where ciamaestro='" & wcia & Right(cod, 3) & "' and status<>'*'"
End If
If (fAbrRst(RSdetermina_mae, Sql$)) Then
   If RSdetermina_mae!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "' ORDER BY cod_maestro3"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "' ORDER BY cod_maestro3"
   End If
End If
If RSdetermina_mae.State = 1 Then RSdetermina_mae.Close
Determina_Maestro_2 = xciamae
End Function
Public Function Numero_Hijos(Codigo, menor, estudio, fechahoy, edad As Integer) As Integer
Dim RSHijos As ADODB.Recordset
Dim manos As Integer
manos = 0
Numero_Hijos = 0
Sql$ = "select * from pladerechohab where cia='" & wcia & "' and placod='" & Codigo & "' and escolar='S' and codvinculo='01' and status<>'*'"
If (fAbrRst(RSHijos, Sql$)) Then
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

Public Function Name_Month(Mes As String) As String
Select Case Mes
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
Public Function DestinoPort(ByVal wSalida As String, ByVal wFileSal As String)

Dim x As Printer
For Each x In Printers
    If x.DeviceName = wSalida Then
       Set Printer = x
    Exit For
    End If
Next

If wImprimeBat = "S" Then
   If Left(Trim(x.DeviceName), 2) <> "\\" Then
       Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & "LPT1", 0)
   Else
       Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & Trim(x.DeviceName), 0)
   End If
Else
    If Left(Printer.Port, 2) = "Ne" Then 'Windows XP
            Dim pos As Integer
            pos = InStr(1, wSalida, "\")
            If pos = 0 Then wSalida = "\\" & GetIPAddress & "\" & wSalida
            CopyFile wFileSal, wSalida, FILE_NOTIFY_CHANGE_LAST_WRITE
    Else ' Windows 9X
       MsgBox "3"
       CopyFile wFileSal, Printer.Port, FILE_NOTIFY_CHANGE_LAST_WRITE
    End If
End If

End Function
'Public Function DestinoPort(ByVal wSalida As String, ByVal wFileSal As String, mshell As Boolean)
'Dim x As Printer
'For Each x In Printers
'    If x.DeviceName = wSalida Then
'       Set Printer = x
'    Exit For
'    End If
'Next
'
'
'Public Function DestinoPort(ByVal wSalida As String, ByVal wFileSal As String)
'
'Dim x As Printer
'For Each x In Printers
'    If x.DeviceName = wSalida Then
'       Set Printer = x
'    Exit For
'    End If
'Next
'
''CopyFile wFileSal, Printer.Port, FILE_NOTIFY_CHANGE_LAST_WRITE
'    If Left(Printer.Port, 2) = "Ne" Then 'Windows XP
'        CopyFile wFileSal, wSalida, FILE_NOTIFY_CHANGE_LAST_WRITE
'    Else ' Windows 9X
'        CopyFile wFileSal, Printer.Port, FILE_NOTIFY_CHANGE_LAST_WRITE
'    End If
'
'End Function
''*********CODIGO PARA RODA*************************************************
''If mshell = True Then
''   Call Shell(App.path & "\Logs\impresora.BAT " & wFileSal & " " & x.DeviceName, 0)
''Else
''   CopyFile wFileSal, Printer.Port, FILE_NOTIFY_CHANGE_LAST_WRITE
''End If
''**************************************************************************
'
''*********CODIGO PARA PPM**************************************************
''If Left(Trim(x.DeviceName), 2) = "EP" Then
''    Call Shell(App.path & "\Logs\impresora.BAT " & wFileSal & " " & "LPT1", 0)
''Else
''   'Call Shell(App.path & "\Logs\impresora.BAT " & wFileSal & " " & "\\EDITH\Epson FX-1180", 0)
''   Call Shell(App.path & "\Logs\impresora.BAT " & wFileSal & " " & Trim(x.DeviceName), 0)
''End If
''**************************************************************************
'
'If wGrupoPla = "01" Then
'    If Left(Printer.Port, 2) = "Ne" Then 'Windows XP
'        CopyFile wFileSal, wSalida, FILE_NOTIFY_CHANGE_LAST_WRITE
'    Else ' Windows 9X
'        CopyFile wFileSal, Printer.Port, FILE_NOTIFY_CHANGE_LAST_WRITE
'    End If
'Else
'
'   If Left(Trim(x.DeviceName), 2) <> "\\" Then
'       Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & "LPT1", 0)
'   Else
'       Call Shell(App.Path & "\Logs\impresora.BAT " & wFileSal & " " & Trim(x.DeviceName), 0)
'   End If
'End If
'End Function

Public Function Ultimo_MES(Mes As Integer) As Integer
Select Case Mes
       Case Is = 1
            Ultimo_MES = 12
End Select
End Function
Public Function m_LetraColumna(ByVal t_nColumna As Integer) As String
    Dim l_sCadena As String
    
    If t_nColumna <= 26 Then
        l_sCadena = Chr$(t_nColumna + 64)
    Else
        If (t_nColumna Mod 26) = 0 Then
            l_sCadena = Chr$((Int((t_nColumna / 26)) - 1) + 64) & "Z"
        Else
            l_sCadena = Chr$(Int((t_nColumna / 26)) + 64) & Chr$((t_nColumna Mod 26) + 64)
        End If
    End If
    m_LetraColumna = l_sCadena
End Function

Public Function Ultimo_Dia(Mes As Integer, ano As Integer) As Integer
Select Case Mes
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
    Formimpri.Lblver.Caption = ver
    Load Formimpri
    Formimpri.Show
    Formimpri.ZOrder 0
End Function

Public Function Remun_Horas(Codigo As String) As Integer
Select Case Codigo
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
       Case Is = "12"
           Remun_Horas = 3
       Case Is = "19"
            Remun_Horas = 4
       Case Is = "21"
            Remun_Horas = 17
       Case Is = "22"
            Remun_Horas = 20
       Case Is = "24"
            Remun_Horas = 18
       Case Is = "25"
            Remun_Horas = 19
       Case Is = "29"
            Remun_Horas = 15
       Case Is = "32"
           Remun_Horas = 5
       Case Is = "43"
           Remun_Horas = 6
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
Public Function fc_ValDecimal(Key As Integer) As String
    Select Case Key
        Case 0, 8, 13, 27, 32, 45, 46, 48 To 57
            Key = Key
        Case Else
            Key = 0
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
Sql$ = "select * from users where sistema='" & wCodSystem & "' and status<>'*' and cod_cia='" & wcia & "' and name_user='" & wuser & "'"
If (fAbrRst(RStp, Sql)) Then
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
         x = 0
         x = mFactProm - Val(Mid(fecproc, 4, 2))
         If x < 0 Then
            fprom = "01/"
            fprom = fprom & Format((x + (12 * masprom)), "00") & "/"
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

Public Sub CargaCombo2(ByRef pCombo As Object, Optional pCodMaestro As String, Optional pSql As String, Optional pTodos As Boolean = False, Optional pNinguno As Boolean = False, Optional pAutoCorrela As Boolean = False)
Dim sSQL As String
Dim Rs As ADODB.Recordset
Dim correla As String

If Len(Trim(pCodMaestro)) > 0 Then
    sSQL = " select cod_maestro2,descrip from maestros_2 where ciamaestro='" & Trim(CODCIA) & (pCodMaestro) & "'"
Else
    sSQL = pSql
End If

If Trim(sSQL) <> "" Then
    Set Rs = OpenRecordset(sSQL, cn)
    

    If Rs.Fields.count > 0 Then
        pCombo.Clear
        Do While Not Rs.EOF
            correla = Right(Rs(0), 4)
            
            If pAutoCorrela Then
                pCombo.AddItem Trim(Rs(1)), , Val(correla)
            Else
                pCombo.AddItem Trim(Rs(1)), , Val(Rs(0))
            End If
            
            Rs.MoveNext
        Loop
            
        Rs.Close
    End If
End If

If pNinguno Then
    pCombo.AddItem Trim("<Ninguno>"), , 0
End If

If pTodos Then
    pCombo.AddItem Trim("<Todos>"), , -1
End If

Set Rs = Nothing

End Sub
Public Sub CargaCombo(ByRef pCombo As Object, Optional pCodMaestro As String, Optional pSql As String, Optional pTodos As Boolean = False, Optional pNinguno As Boolean = False, Optional pAutoCorrela As Boolean = False)
Dim sSQL As String
Dim Rs As Object
Dim correla As String

If Len(Trim(pCodMaestro)) > 0 Then
    sSQL = " select cod_maestro2,descrip from maestros_2 where ciamaestro='" & Trim(CODCIA) & (pCodMaestro) & "'"
Else
    sSQL = pSql
End If

If Trim(sSQL) <> "" Then
Set Rs = cn.Execute(sSQL)

    If Rs.Fields.count > 0 Then
        pCombo.Clear
        Do While Not Rs.EOF
            correla = Right(Rs(0), 4)
            
            If pAutoCorrela Then
                pCombo.AddItem Trim(Rs(1)), , Val(correla)
            Else
                pCombo.AddItem Trim(Rs(1)), , Val(Rs(0))
            End If
            
            Rs.MoveNext
        Loop
            
        Rs.Close
    End If
End If

If pNinguno Then
    pCombo.AddItem Trim("<Ninguno>"), , 0
End If

If pTodos Then
    pCombo.AddItem Trim("<Todos>"), , 999
End If

Set Rs = Nothing

End Sub

Public Function Semana_Calcular(ByVal pSemana As Integer, ByVal pSemanaPago As Integer, _
AñoProceso As String, CodEmpresa As String) As Boolean

'************codigo nuevo agregado giovanni 13092007***************************
Dim rs_Semana_Calculo As ADODB.Recordset
Dim s_Semana_Calculo As String

s_Semana_Calculo = "select Calculo_Asigfam from plasemanas where cia='" & CodEmpresa & "' " & _
"and status <> '*' and ano='" & AñoProceso & "' and semana=" & pSemana & " "
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
Public Function OpenRecordset(ByVal Strsql As String, cn As ADODB.Connection) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    With Rs
        .CursorLocation = adUseClient
        .Open Strsql, cn, adOpenDynamic, adLockOptimistic
    End With
    Set OpenRecordset = Rs
    Set Rs = Nothing
End Function
Public Sub Centrar(frmNombre As Form)
    With frmNombre
        If frmNombre.MDIChild = True Then
            .Left = (MDIplared.ScaleWidth - .Width) / 2
            .Top = (MDIplared.ScaleHeight - .Height) / 2
        Else
            .Left = (Screen.Width - .ScaleWidth) / 2
            .Top = (Screen.Height - .ScaleHeight) / 2
        End If
    End With
End Sub

Sub LeerGrillaClasica(Bookmark As Variant, Col As Integer, Value As Variant, MAXROW As Long, MAXCOL As Integer, GridArray As Variant)
    Dim Index As Long
    Index = IndexFromBookmark(Bookmark, 0, MAXROW)
    If Index < 0 Or Index >= MAXROW Or Col < 0 Or _
       Col > MAXCOL Then
        Value = Null
    Else
        Value = GridArray(Col, Index)
    End If
End Sub
Sub EscribirGrillaClasica(Bookmark As Variant, Col As Integer, Value As Variant, MAXROW As Long, MAXCOL As Integer, GridArray() As Variant, Tdbgrid As Object)
    If Not StoreUserData(Bookmark, Col, Value, MAXROW _
    , MAXCOL, GridArray(), Tdbgrid) Then Bookmark = Null
End Sub
Sub AgregarGrillaClasica(NewRowBookmark As Variant, Col As Integer, Value As Variant, TDBGrid1 As Object, MAXCOL As Integer, MAXROW As Long, GridArray() As Variant)
    If IsNull(NewRowBookmark) Then
       NewRowBookmark = GetNewBookmark(TDBGrid1, MAXCOL, MAXROW, GridArray())
    End If
    If Not StoreUserData(NewRowBookmark, Col, Value, MAXROW _
    , MAXCOL, GridArray(), TDBGrid1) Then
        DeleteRow NewRowBookmark, TDBGrid1, MAXCOL, MAXROW, GridArray()
        NewRowBookmark = Null
    End If
End Sub
Sub BorrarGrillaClasica(Bookmark As Variant, TDBGrid1 As Object, MAXCOL As Integer, MAXROW As Long, GridArray() As Variant)
    If Not DeleteRow(Bookmark, TDBGrid1, MAXCOL, MAXROW, GridArray()) Then Bookmark = Null
End Sub

Function StoreUserData(Bookmark As Variant, _
        Col As Integer, Userval As Variant, MAXROW As Long _
        , MAXCOL As Integer, GridArray() As Variant, Optional TDBGrid1 As Object) As Boolean
    Dim Index As Long
    
      Index = IndexFromBookmark(Bookmark, 0, MAXROW)
    If Index < 0 Or Index >= MAXROW Or Col < 0 Or _
       Col > MAXCOL Then
        StoreUserData = False
    Else
        StoreUserData = True
        
        GridArray(Col, Index) = Userval
    End If
End Function

Function IndexFromBookmark(Bookmark As Variant, offset As Long, MAXROW As Long) As Long
    Dim Index As Long
    If IsNull(Bookmark) Then
        If offset < 0 Then
            Index = MAXROW + offset
        Else
            Index = -1 + offset
        End If
    Else
        Index = Val(Bookmark) + offset
    End If
    If Index >= 0 And Index < MAXROW Then
       IndexFromBookmark = Index
    Else
       IndexFromBookmark = -9999
    End If
End Function
Function GetNewBookmark(TDBGrid1 As Object, MAXCOL As Integer, MAXROW As Long, GridArray() As Variant) As Variant
    ' Es llamada cuando se necesita crear un bookmark para una nueva fila
    ReDim Preserve GridArray(0 To MAXCOL, 0 To MAXROW)
    GetNewBookmark = Str$(MAXROW)
    MAXROW = MAXROW + 1
    TDBGrid1.ApproxCount = MAXROW
End Function

Function DeleteRow(Bookmark As Variant, TDBGrid1 As Object, _
MAXCOL As Integer, MAXROW As Long, GridArray() As Variant) As Boolean
    ' Esta función es llamada para eleminar lógicamente un registro
    ' del arreglo. La fila a ser eliminada está dada por el Parámetro
    ' bookmark
    Dim I As Long, Index As Long
    Dim J As Integer

    Index = IndexFromBookmark(Bookmark, 0, MAXROW)
    
    If Index < 0 Or Index >= MAXROW Then
        'El indice de la fila no es válido
        DeleteRow = False
        Exit Function
    End If
    MAXROW = MAXROW - 1
    For I = Index To MAXROW - 1
        For J = 0 To MAXCOL - 1
            GridArray(J, I) = GridArray(J, I + 1)
        Next J
    Next I
    If MAXROW > 0 Then
        ReDim Preserve GridArray(0 To MAXCOL, 0 _
              To MAXROW - 1)
    Else
        ReDim GridArray(0 To MAXCOL, 0)
    End If
    DeleteRow = True
    ' Calibra el scroll bar basado en el nuevo tamaño de los datos
    TDBGrid1.ApproxCount = MAXROW
End Function

Public Function Ejecuta(ByVal pSql As String) As Boolean

If Len(Trim(pSql)) = 0 Then
    Strerror = "Ingresar la Sentencia a Ejecutar"
    Exit Function
End If

If cn.State = StateClosed Then
    Strerror = "Conexion Cerrada"
    Exit Function
End If

On Error GoTo Ejecuta
    
    cn.Execute pSql

    Ejecuta = True
    
Exit Function

Ejecuta:
Strerror = Err.Description

End Function

Rem  ADD 2011.04.20 ======================================================================================================

Public Function Trae_CIA(ByVal pCia As String, Optional ByVal pTipo As Integer = 0) As String
    Rem Editado el 2011.12.21
    Dim Rs As ADODB.Recordset
    Dim Cadena As String
    
    Cadena = "SP_TRAE_NOM_COMERCIAL '" & pCia & "', " & pTipo & ""
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open Cadena, cn, adOpenStatic, adLockOptimistic
    
    If Not Rs.EOF Then
        Trae_CIA = UCase(Trim(Rs(0)))
    Else
        Cadena = "SELECT RAZSOC FROM CIA WHERE COD_CIA = '" & pCia & "' AND STATUS != '*'"
        If Rs.State = adStateOpen Then Rs.Close
        Set Rs = OpenRecordset(Cadena, cn)
        If Not Rs.EOF Then Trae_CIA = UCase(Trim(Rs(0))) Else Trae_CIA = Empty
    End If
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
        Set Rs = Nothing
    End If
End Function

Public Function Trae_RUC(ByVal cia As String) As String
    Dim Rs As ADODB.Recordset
    Dim Cadena As String
    Cadena = "SELECT RUC FROM CIA WHERE COD_CIA='" & cia & "'"
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open Cadena, cn, adOpenStatic, adLockOptimistic
    If Not Rs.EOF Then
        Trae_RUC = UCase(Trim(Rs!RUC))
    Else
        Trae_RUC = Empty
    End If
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Clone
        Set Rs = Nothing
    End If
End Function

Public Function Trae_DIRECCION(ByVal cia As String) As String
    Dim Rs As ADODB.Recordset
    Dim Cadena As String
    'Cadena = "SELECT RTRIM(DIRECC) + ' # ' + RTRIM(NRO) + ' DPTO N°' + RTRIM(DPTO) AS DIRECC FROM CIA WHERE COD_CIA='" & cia & "'"
    Cadena = "SP_TRAE_DIRECCION_CIA '" & cia & "'"
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open Cadena, cn, adOpenStatic, adLockOptimistic
    If Not Rs.EOF Then
        Trae_DIRECCION = UCase(Trim(Rs!direcc))
    Else
        Trae_DIRECCION = Empty
    End If
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Clone
        Set Rs = Nothing
    End If
End Function

Public Function EXEC_SQL(ByVal mSQL As String, ByVal mConnection As ADODB.Connection) As Boolean
    EXEC_SQL = False
    Dim InTrans As Boolean
    On Error GoTo MyErr
    mConnection.CommandTimeout = 0
    mConnection.BeginTrans: InTrans = True
    mConnection.Execute (mSQL)
    mConnection.CommitTrans: InTrans = False
    EXEC_SQL = True
MyErr:
    If InTrans Then mConnection.RollbackTrans
    If Err.Number <> 0 Then
        'MsgBox Err.Description, vbExclamation, Err.Source
        Err.Clear
    End If
End Function

Public Sub Up(Ctrl As Object)
    If Trim(Ctrl.Text) = "" Then
       Ctrl.Text = Format(Year(Date), "0000")
    Else
       Ctrl.Text = Ctrl.Text + 1
    End If
End Sub

Public Sub Down(Ctrl As Object)
    If Trim(Ctrl.Text) = "" Then
       Ctrl.Text = Format(Year(Date), "0000")
    Else
       Ctrl.Text = Ctrl.Text - 1
    End If
End Sub

Public Sub NumberOnly(ByRef Key As Integer)
    If InStr("0123456789" & Chr(8), Chr(Key)) = 0 Then Key = 0
End Sub

Public Function mFormato_Fecha(ByVal mDate As Date) As String
    mFormato_Fecha = Empty
    mFormato_Fecha = Format(mDate, "yyyyMMdd")
End Function

Public Sub Trae_Tipo_Trab(Ctrl As Object)
    Cadena = "SP_TRAE_TIPO_TRAB"
    Call rCarCbo(Ctrl, Cadena, "XX", "00")
    Ctrl.ListIndex = 0
End Sub

Public Sub Trae_Tipo_Boleta(Ctrl As Object, tipo As Integer)
    ' Tipo=1: TODOS
    ' Tipo=2: TODOS MENOS DEPOSITO CTS Y TRANSFERENCIAS
    Cadena = "usp_Pla_ConsultarTipoBoleta " & Val(tipo)
    Call rCarCbo(Ctrl, Cadena, "XX", "00")
    Ctrl.ListIndex = 0
End Sub





Function Trae_Meses_Laborados(ByVal cia As String, ByVal PlaCod As String, ByVal FechaProceso As Date) As Integer
    Dim Rs As ADODB.Recordset
    Cadena = "SP_TRAE_MESES_LABORADOS " & _
            "'" & wcia & "', " & _
            "'" & PlaCod & "', " & _
            "'" & FechaProceso & "'"
    Set Rs = OpenRecordset(Cadena, cn)
    Trae_Meses_Laborados = Rs(0)
    If Rs.State = adStateOpen Then Rs.Close
    Set Rs = Nothing
End Function

Public Function Trae_Porc_EsSalud_EPS(ByVal PlaCod As String, ByVal T_Trab As String) As Double
        Dim MCADENA     As String
        Dim Afecto_EPS  As Boolean
        Dim rsTemporal  As ADODB.Recordset
        Dim FactorEPS As Double
        Dim CodEPS As String
        
        'MCADENA = "SELECT ISNULL(APORTACION_EPS,0) AS APORTACION_EPS FROM CIA WHERE COD_CIA = '" & wcia & "' AND STATUS != '*'"
        
        MCADENA = "EXEC spDevolverAfilEPS '" & wcia & "','" & PlaCod & "'"
        
        Set rsTemporal = OpenRecordset(MCADENA, cn)
        If Not rsTemporal.EOF Then
            Afecto_EPS = rsTemporal!afiliado_eps_serv
            CodEPS = rsTemporal!codigo_eps
        Else
            Afecto_EPS = 0
        End If
        
        If rsTemporal.State = adStateOpen Then rsTemporal.Close
        If Afecto_EPS Then
            If CodEPS <> "" Then
                MCADENA = "SELECT ISNULL(plaimporte,0) AS IMPORTE FROM MAESTROS_2 WHERE RIGHT(CIAMAESTRO,3) = '143' AND STATUS != '*' AND RTRIM(ISNULL(CODSUNAT,'')) != '' AND COD_MAESTRO2 = '" & CodEPS & "'"
                If rsTemporal.State = adStateOpen Then rsTemporal.Close
                Set rsTemporal = OpenRecordset(MCADENA, cn)
                If Not rsTemporal.EOF Then
                    Trae_Porc_EsSalud_EPS = Val(rsTemporal!importe)
                Else
                    Trae_Porc_EsSalud_EPS = 0
                End If
            Else
                Trae_Porc_EsSalud_EPS = 0
            End If
        Else
            MCADENA = "SELECT APORTACION FROM PLACONSTANTE WHERE TIPOMOVIMIENTO='03' AND CODINTERNO='01' AND STATUS!='*' AND CIA='" & wcia & "'"
            If (fAbrRst(rsTemporal, MCADENA)) Then
                Trae_Porc_EsSalud_EPS = rsTemporal(0)
            End If
        End If
End Function

Public Function Exportar_Excel(ByRef mRecordSet As ADODB.Recordset) As Boolean
    
    On Error GoTo errSub
    
    Dim cn          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    Set hoja = Libro.Worksheets(1)
    Excel.Visible = True: Excel.UserControl = True
    iCol = mRecordSet.Fields.count
    For iCol = 1 To mRecordSet.Fields.count
        hoja.Cells(1, iCol).Value = mRecordSet.Fields(iCol - 1).Name
    Next
    
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 11 Then
        hoja.Cells(2, 1).CopyFromRecordset mRecordSet
    Else
        arrData = mRecordSet.GetRows
        iRec = UBound(arrData, 2) + 1
        For iCol = 0 To mRecordSet.Fields.count - 1
            For iRow = 0 To iRec - 1
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))

                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
        hoja.Cells(2, 1).Resize(iRec, mRecordSet.Fields.count).Value = GetData(arrData)
    End If

    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit
    
    If Not mRecordSet Is Nothing Then
        If mRecordSet.State = adStateOpen Then mRecordSet.Close
        Set mRecordSet = Nothing
    End If

    Set hoja = Nothing
    Set Libro = Nothing
    Excel.Visible = True
    
    Exportar_Excel = True
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_Excel = False
End Function

Public Function GetData(vValue As Variant) As Variant
    Dim x As Long, Y As Long, xMax As Long, yMax As Long, t As Variant
    
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
    
    ReDim t(xMax, yMax)
    For x = 0 To xMax
        For Y = 0 To yMax
            t(x, Y) = vValue(Y, x)
        Next Y
    Next x
    GetData = t
End Function

'IMPLEMENTACION GALLOS MARMOLERIA

Public Sub LimpiarRsT(ByRef pRs As ADODB.Recordset, ByRef pDgrd As TrueOleDBGrid70.Tdbgrid)
pDgrd.Refresh
Set pDgrd.DataSource = Nothing
If pRs.State = 1 Then
    If pRs.RecordCount > 0 Then
        pRs.MoveFirst
        Do While Not pRs.EOF
            pRs.Delete
            If Not pRs.EOF Then pRs.MoveNext
        Loop
    End If
End If
Set pDgrd.DataSource = pRs
pDgrd.Refresh
End Sub
Public Sub Carga_Asiento_Excel(ano As Integer, Mes As Integer, lote As String, Voucher As String, titulo As String, MesDes As String, Resumen As String)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object



Sql = "spListaAsientoGral '" & wcia & "'," & ano & "," & Mes & ",'" & lote & "','" & Voucher & "'"
If wGrupoPla = "01" And wcia = "21" Then
   cn21.Open
   Set Rs = cn21.Execute(Sql)
   If Not Rs.EOF Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub
Else
   If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub
End If


Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:D").ColumnWidth = 12
xlSheet.Range("E:E").ColumnWidth = 23
xlSheet.Range("E:E").HorizontalAlignment = xlCenter
xlSheet.Range("C:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 1).Value = NOMBREEMPRESA
xlSheet.Cells(3, 1).Value = titulo & " " & lote & " - " & Voucher
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 5)).Merge
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 5)).HorizontalAlignment = xlCenter
xlSheet.Cells(5, 1).Value = MesDes & " - " & Str(ano)
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 5)).Merge
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 5)).HorizontalAlignment = xlCenter
xlSheet.Cells(7, 1).Value = "CUENTA"
xlSheet.Cells(7, 2).Value = "AUXILIAR"
xlSheet.Cells(7, 3).Value = "DEBE"
xlSheet.Cells(7, 4).Value = "HABER"
xlSheet.Cells(7, 5).Value = "REFERENCIA"
xlSheet.Cells(7, 6).Value = "REFERENCIA2"
xlSheet.Cells(7, 7).Value = "ANEXO"
xlSheet.Cells(7, 8).Value = "ORDEN"
xlSheet.Range(xlSheet.Cells(7, 1), xlSheet.Cells(7, 5)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(7, 1), xlSheet.Cells(7, 5)).HorizontalAlignment = xlCenter

nFil = 8
Dim tDebe As Currency
Dim tHAber As Currency
tDebe = 0: tHAber = 0
Do While Not Rs.EOF
   If Rs!tipo = "D" Then
        xlSheet.Cells(nFil, 1).Value = "'" & Trim(Rs!CGCOD)
        xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!cgaux)
        If Trim(Rs!CGMOV) = "1" Then
           xlSheet.Cells(nFil, 3).Value = Rs!cgimporte
           tDebe = tDebe + Rs!cgimporte
        Else
           xlSheet.Cells(nFil, 4).Value = Rs!cgimporte
           tHAber = tHAber + Rs!cgimporte
        End If
        xlSheet.Cells(nFil, 5).Value = Trim(Rs!cgref)
        xlSheet.Cells(nFil, 6).Value = Trim(Rs!refe2)
        xlSheet.Cells(nFil, 7).Value = Trim(Rs!anexo2)
        xlSheet.Cells(nFil, 8).Value = Trim(Rs!Orden)
        nFil = nFil + 1
   End If
   Rs.MoveNext
Loop
xlSheet.Cells(nFil, 1).Value = "TOTAL"
xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 2)).Merge
xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 2)).HorizontalAlignment = xlCenter
xlSheet.Cells(nFil, 3).Value = tDebe
xlSheet.Cells(nFil, 4).Value = tHAber
xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 4)).Borders.LineStyle = xlContinuous

If Resumen = "S" Then
    nFil = nFil + 2
    xlSheet.Cells(nFil, 1).Value = "RESUMEN"
    nFil = nFil + 2
    
    xlSheet.Cells(nFil, 1).Value = "CUENTA"
    xlSheet.Cells(nFil, 2).Value = "AUXILIAR"
    xlSheet.Cells(nFil, 3).Value = "DEBE"
    xlSheet.Cells(nFil, 4).Value = "HABER"
    xlSheet.Range(xlSheet.Cells(7, 1), xlSheet.Cells(7, 4)).Borders.LineStyle = xlContinuous
    xlSheet.Range(xlSheet.Cells(7, 1), xlSheet.Cells(7, 4)).HorizontalAlignment = xlCenter
    nFil = nFil + 2
    If Rs.RecordCount > o Then Rs.MoveFirst
    tDebe = 0: tHAber = 0
    Do While Not Rs.EOF
       If Rs!tipo = "R" Then
            xlSheet.Cells(nFil, 1).Value = "'" & Trim(Rs!CGCOD)
            xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!cgaux)
            If Trim(Rs!CGMOV) = "1" Then
               xlSheet.Cells(nFil, 3).Value = Rs!cgimporte
               tDebe = tDebe + Rs!cgimporte
            Else
               xlSheet.Cells(nFil, 4).Value = Rs!cgimporte
               tHAber = tHAber + Rs!cgimporte
            End If
            xlSheet.Cells(nFil, 5).Value = Trim(Rs!cgref)
            nFil = nFil + 1
       End If
       Rs.MoveNext
    Loop
    xlSheet.Cells(nFil, 1).Value = "TOTAL"
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 2)).Merge
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 2)).HorizontalAlignment = xlCenter
    xlSheet.Cells(nFil, 3).Value = tDebe
    xlSheet.Cells(nFil, 4).Value = tHAber
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 4)).Borders.LineStyle = xlContinuous
End If
Rs.Close: Set Rs = Nothing



xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
'xlApp2.Application.Caption = "REPORTE DE PERSONAL"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

If wGrupoPla = "01" And wcia = "21" Then cn21.Close

Screen.MousePointer = vbDefault
End Sub
Public Function NamePC() As String
NamePC = ""
Dim nPC As String
Dim buffer As String
Dim estado As Long
buffer = String$(255, " ")
estado = GetComputerName(buffer, 255)
If estado <> 0 Then
nPC = Left(buffer, 255)
End If
NamePC = nPC
End Function

Public Sub Memo_Cese(lCod As String)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object

Sql = "uSp_Pla_Memo_Cese '" & wcia & "','" & lCod & "'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub


Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("C:C").HorizontalAlignment = xlLeft
xlSheet.Range("A:A").ColumnWidth = 3
xlSheet.Range("B:B").ColumnWidth = 28
xlSheet.Range("C:C").ColumnWidth = 40
xlSheet.Cells(2, 2).Value = NOMBREEMPRESA
xlSheet.Cells(4, 2).Value = "MEMORANDUM"
xlSheet.Cells(4, 2).Font.Size = 14
xlSheet.Cells(4, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(4, 2), xlSheet.Cells(4, 3)).Merge
xlSheet.Cells(5, 2).Value = "FECHA: " & Format(Day(Date), "00") + " de " + mes_palabras2(Month(Date)) + " del " + Format(Year(Date), "0000")
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 3)).Merge
xlSheet.Cells(5, 2).HorizontalAlignment = xlCenter
xlSheet.Cells(7, 2).Value = "DE:"
xlSheet.Cells(7, 3).Value = "Dpto. Personal"
xlSheet.Cells(8, 2).Value = "PARA:"
xlSheet.Cells(8, 3).Value = "Dpto. Contabilidad"
xlSheet.Cells(10, 2).Value = "DATOS DEL TRABAJADOR A LIQUIDAR"
xlSheet.Range(xlSheet.Cells(10, 2), xlSheet.Cells(10, 3)).Borders(xlEdgeBottom).LineStyle = xlDouble

xlSheet.Cells(12, 2).Value = "CODIGO"
xlSheet.Cells(12, 3).Value = Trim(Rs!PlaCod & "")
xlSheet.Cells(13, 2).Value = "NOMBRE"
xlSheet.Cells(13, 3).Value = Trim(Rs!nombre & "")
xlSheet.Cells(14, 2).Value = "DNI"
xlSheet.Cells(14, 3).Value = Trim(Rs!DNI & "")
xlSheet.Cells(15, 2).Value = "CARGO"
xlSheet.Cells(15, 3).Value = Trim(Rs!Cargo & "")
xlSheet.Cells(16, 2).Value = "REGIMEN PENSIONARIO"
xlSheet.Cells(16, 3).Value = Trim(Rs!PENSION & "")
xlSheet.Cells(17, 2).Value = "SUELDO S/."
xlSheet.Cells(17, 3).Value = Rs!sueldo
xlSheet.Cells(17, 3).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(18, 2).Value = "ASIGNACION FAMILIAR"
xlSheet.Cells(18, 3).Value = Rs!asigfam
xlSheet.Cells(18, 3).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(19, 2).Value = "FECHA INGRESO"
xlSheet.Cells(19, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
xlSheet.Cells(20, 2).Value = "FECHA DE CESE"
xlSheet.Cells(20, 3).Value = "'" & Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000")
xlSheet.Cells(21, 2).Value = "MOTIVO CESE"
xlSheet.Cells(21, 3).Value = Trim(Rs!motivo & "")
xlSheet.Cells(22, 2).Value = "CTA. CTE S/."
xlSheet.Cells(22, 3).Value = Rs!ctacte
xlSheet.Cells(22, 3).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(23, 2).Value = "FECHA ULT. PROV. VACA."
xlSheet.Cells(23, 3).Value = Trim(Rs!PROV & "")
xlSheet.Cells(24, 2).Value = "ULTIMA GRATIFICACION S/."
xlSheet.Cells(24, 3).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(24, 3).Value = Rs!grati
xlSheet.Cells(25, 2).Value = "CENTRO DE COSTO"
xlSheet.Cells(25, 3).Value = Trim(Rs!ccosto & "")
xlSheet.Cells(26, 2).Value = "BANCO DE CTS"
xlSheet.Cells(26, 3).Value = Trim(Rs!CtsBco & "")
xlSheet.Cells(27, 2).Value = "NRO CTA CTS"
xlSheet.Cells(27, 3).Value = Trim(Rs!ctsnumcta & "")
xlSheet.Cells(28, 2).Value = "MONEDA CTS"
xlSheet.Cells(28, 3).Value = Trim(Rs!ctsmoneda & "")
xlSheet.Cells(29, 2).Value = "BANCO HABERES"
xlSheet.Cells(29, 3).Value = Trim(Rs!PagoBco & "")
xlSheet.Cells(30, 2).Value = "CUENTA HABERES"
xlSheet.Cells(30, 3).Value = Trim(Rs!pagonumcta & "")

Rs.Close

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Screen.MousePointer = vbDefault
End Sub
Public Sub Tabla_dinamica(cia As String, ano1 As Integer, ano2 As Integer, TipoTrab As String)
Dim nFil As Integer
Dim nCol As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object



Dim sTableName As String
sTableName = "##TablaDin" + wuser

Sql = "uSp_Tabla_Dinamica '" & wcia & "'," & ano1 & "," & ano2 & ",'" & TipoTrab & "','" & sTableName & "'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Datos"
'xlApp1.ActiveSheet.Select
With xlSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
     "OLEDB;User ID=" & Trim(wuser) & ";PWD=" & Trim(wclave) & ";Data Source=" & wserver & ";Initial Catalog=" & WDatabase & ";Provider=SQLOLEDB.1"), Destination:=xlApp1.ActiveSheet.Range("$A$1")).QueryTable
     .CommandType = xlCmdTable
     .CommandText = Array("" & WDatabase & ".""dbo""." & sTableName & "")
     .AdjustColumnWidth = True
     .Refresh BackgroundQuery:=False
End With

Dim I As Integer
Dim mTipoMov As String
For I = 0 To Rs.Fields.count - 1
    If I > 18 Then
       If UCase(Left(Rs.Fields(I).Name, 1)) = "I" Then
          mTipoMov = "02"
       ElseIf UCase(Left(Rs.Fields(I).Name, 1)) = "D" Or UCase(Left(Rs.Fields(I).Name, 1)) = "A" Then
          mTipoMov = "03"
       ElseIf UCase(Left(Rs.Fields(I).Name, 1)) = "H" Then
          mTipoMov = "**"
       End If
       
'       If I = 102 Then
'         Stop
'       End If
       If (Trim(Rs.Fields(I).Name) = "Ctas" Or Trim(Rs.Fields(I).Name) = "Suspension" Or Trim(Rs.Fields(I).Name) = "Afiliado_al_sindicato" Or Trim(Rs.Fields(I).Name) = "Mes_Año") Then
            xlSheet.Cells(1, I + 1).Value = UCase(Rs.Fields(I).Name)
       Else
            xlSheet.Cells(1, I + 1).Value = UCase(Left(Rs.Fields(I).Name, 1)) & "_" & Devuelve_Nombre_Campo(mTipoMov, Mid(Rs.Fields(I).Name, 2, 2))
       End If
    End If
Next

    
'If xlApp2.Sheets.count < 2 Then xlApp2.Sheets.Add
'Set xlSheet = xlApp2.Worksheets(2)
'xlApp2.Sheets(2).Select
'xlSheet.Name = "Dinamica"
 
On Error GoTo Termina

'ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'    "Tabla_DatosExternos_1", Version:=xlPivotTableVersion12).CreatePivotTable _
'    TableDestination:="Dinamica!F3C1", TableName:="Tabla dinámica1", _
'    DefaultVersion:=xlPivotTableVersion12
'
'Sheets("Dinamica").Select
'Cells(3, 1).Select
Termina:

xlApp2.Application.ActiveWindow.DisplayGridlines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
Public Function Devuelve_Nombre_Campo(mTipo, Codigo) As String
Dim rsDev As ADODB.Recordset
Set rsDev = New ADODB.Recordset

Dim cad As String
If mTipo = "**" Then
   cad = "select descrip from maestros_2 where ciamaestro='01077' and cod_maestro2='" & Codigo & "'"
Else
   cad = "select top 1 descripcion as descrip from placonstante where cia='" & wcia & "' and tipomovimiento='" & mTipo & "' and codinterno='" & Codigo & "' and status<>'*'"
End If
Set rsDev = cn.Execute(cad)
If Not rsDev.BOF And Not rsDev.EOF Then Devuelve_Nombre_Campo = Trim(rsDev!DESCRIP & "")
If rsDev.State = 1 Then rsDev.Close
End Function
Public Sub Marcaciones(cia As String, f1 As String, f2 As String, TipoTrab As String)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object

Sql = "uSp_Traer_Marcaciones '" & cia & "','" & f1 & "','" & f2 & "','" & TipoTrab & "'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Marcaciones"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 8
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").ColumnWidth = 9.5
xlSheet.Range("D:Z").HorizontalAlignment = xlCenter
xlSheet.Range("B:Z").NumberFormat = "@"

xlSheet.Cells(1, 1).Value = NOMBREEMPRESA

xlSheet.Cells(3, 2).Value = "CODIGO"
xlSheet.Cells(3, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 3)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 3)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(4, Rs.Fields.count + 1)).Interior.ColorIndex = 37
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(4, Rs.Fields.count + 1)).Interior.Pattern = xlSolid
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(4, Rs.Fields.count + 1)).Font.Bold = True

Dim I As Integer
Dim lName As String
For I = 3 To Rs.Fields.count - 1
    If CInt(I) Mod 2 <> 0 Then
       lName = Dia_Sem(Rs(I).Name)
       lName = lName + " " + Mid(Rs(I).Name, 8, 2) + " de " & Mid(Name_Month(Mid(Rs(I).Name, 5, 2)), 1, 3)
       xlSheet.Cells(3, I + 1).Value = lName
       xlSheet.Range(xlSheet.Cells(3, I + 1), xlSheet.Cells(3, I + 2)).Merge
       xlSheet.Cells(4, I + 1).Value = "ENTRADA"
    Else
       xlSheet.Cells(4, I + 1).Value = "SALIDA"
    End If
Next
xlSheet.Cells(4, I + 1).Value = "SALIDA"

nFil = 5
Dim lCount As Integer
lCount = 1
Do While Not Rs.EOF
   If Mid(Rs!Codigo, 1, 2) <> "**" Then
      For I = 0 To Rs.Fields.count - 1
         xlSheet.Cells(nFil, 1).Value = lCount
         If I > 1 Then
            xlSheet.Cells(nFil, I + 2).Value = Mid(Rs(I), 1, 5)
         Else
            xlSheet.Cells(nFil, I + 2).Value = Trim(Rs(I) & "")
         End If
      Next
      nFil = nFil + 1
      lCount = lCount + 1
   End If
   Rs.MoveNext
Loop

xlSheet.Range("B3:B4").Merge
xlSheet.Range("C3:C4").Merge

xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(nFil - 1, Rs.Fields.count + 1)).Borders.LineStyle = xlContinuous

nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "Observaciones"
nFil = nFil + 2
If Rs.RecordCount > 0 Then Rs.MoveFirst
Dim lMarca As String
Do While Not Rs.EOF
   If Mid(Rs!Codigo, 1, 2) = "**" Then
       xlSheet.Cells(nFil, 2).Value = Mid(Rs!Codigo, 3, 10)
       xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
       xlSheet.Cells(nFil, 4).Value = "Trabajador sin marcaciones"
       xlSheet.Cells(nFil, 4).HorizontalAlignment = xlLeft
       nFil = nFil + 1
   Else
      For I = 0 To Rs.Fields.count - 1
         If I > 1 Then
            lName = Dia_Sem(Rs(I).Name)
            lName = lName + " " + Mid(Rs(I).Name, 8, 2) + " de " & Mid(Name_Month(Mid(Rs(I).Name, 5, 2)), 1, 3)
            
            If CInt(I) Mod 2 = 0 Then
               lMarca = Mid(Rs(I), 1, 5)
            Else
               If Trim(Rs(I) & "") = "" Then
                  If lMarca = "" Then
                     xlSheet.Cells(nFil, 2).Value = Trim(Rs!Codigo & "")
                     xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
                     xlSheet.Cells(nFil, 4).Value = Mid(Rs(I).Name, 8, 2) & "/" & Mid(Rs(I).Name, 5, 2) & "/20" & Mid(Rs(I).Name, 2, 2)
                     xlSheet.Cells(nFil, 5).Value = "FALTO"
                     xlSheet.Cells(nFil, 4).HorizontalAlignment = xlLeft
                     xlSheet.Cells(nFil, 5).HorizontalAlignment = xlLeft
                     nFil = nFil + 1
                  Else
                     xlSheet.Cells(nFil, 2).Value = Trim(Rs!Codigo & "")
                     xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
                     xlSheet.Cells(nFil, 4).Value = Mid(Rs(I).Name, 8, 2) & "/" & Mid(Rs(I).Name, 5, 2) & "/20" & Mid(Rs(I).Name, 2, 2)
                     xlSheet.Cells(nFil, 5).Value = "Solo Registra una marcación   " & lMarca
                     xlSheet.Cells(nFil, 4).HorizontalAlignment = xlLeft
                     xlSheet.Cells(nFil, 5).HorizontalAlignment = xlLeft
                     nFil = nFil + 1
                  End If
               ElseIf lMarca = "" Then
                  xlSheet.Cells(nFil, 2).Value = Trim(Rs!Codigo & "")
                  xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
                  xlSheet.Cells(nFil, 4).Value = Mid(Rs(I).Name, 8, 2) & "/" & Mid(Rs(I).Name, 5, 2) & "/20" & Mid(Rs(I).Name, 2, 2)
                  xlSheet.Cells(nFil, 5).Value = "Solo Registra una marcación   " & Mid(Rs(I), 1, 5)
                  xlSheet.Cells(nFil, 4).HorizontalAlignment = xlLeft
                  xlSheet.Cells(nFil, 5).HorizontalAlignment = xlLeft
                  nFil = nFil + 1
               End If
               lMarca = ""
            End If
         End If
         
         
         
      Next
   End If
   Rs.MoveNext
Loop

Rs.Close: Set Rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridlines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
Private Function Dia_Sem(lDay As String) As String
       Select Case UCase(Mid(lDay, 11, 3))
       Case "MON": Dia_Sem = "Lunes"
       Case "TUE": Dia_Sem = "Martes"
       Case "WED": Dia_Sem = "Miercoles"
       Case "THU": Dia_Sem = "Jueves"
       Case "FRI": Dia_Sem = "Viernes"
       Case "SAT": Dia_Sem = "Sabado"
       Case "SUN": Dia_Sem = "Domingo"
       End Select
End Function
Public Function Llenar_Ceros(lText As String, lNum As Integer) As String
Llenar_Ceros = lText
Dim I As Integer
For I = Len(lText) To lNum - 1
    Llenar_Ceros = "0" + Llenar_Ceros
Next
End Function

'ADD LFSA - 26/10/2012 VALIDA NUMERO DECIMAL
Public Function fc_ValidarDecimal(ByVal KeyAscii As Integer, ByVal texto As String) As Integer
        If Chr(KeyAscii) = "." Then

            If Len(texto) = 0 Then
                If Chr(KeyAscii) = "." Then
                    KeyAscii = 0
                    fc_ValidarDecimal = KeyAscii
                    Exit Function
                End If
            End If

            If InStr(1, texto, ".") <> 0 Then
                KeyAscii = 0
                fc_ValidarDecimal = KeyAscii
                Exit Function
            End If
        End If

        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
            fc_ValidarDecimal = KeyAscii
            Exit Function
        End If
        fc_ValidarDecimal = KeyAscii
    End Function
Public Sub Tardanzas(cia As String, f1 As String, f2 As String, TipoTrab As String)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object


Sql = "uSp_Traer_Marcaciones '" & cia & "','" & f1 & "','" & f2 & "','" & TipoTrab & "'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Marcaciones"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 8
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").ColumnWidth = 9.5
xlSheet.Range("D:Z").HorizontalAlignment = xlCenter
xlSheet.Range("B:Z").NumberFormat = "@"

xlSheet.Cells(1, 1).Value = NOMBREEMPRESA

xlSheet.Cells(3, 2).Value = "CODIGO"
xlSheet.Cells(3, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 3)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 3)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(4, Rs.Fields.count + 1)).Interior.ColorIndex = 37
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(4, Rs.Fields.count + 1)).Interior.Pattern = xlSolid
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(4, Rs.Fields.count + 1)).Font.Bold = True

Dim I As Integer
Dim lName As String
For I = 3 To Rs.Fields.count - 1
    If CInt(I) Mod 2 <> 0 Then
       lName = Dia_Sem(Rs(I).Name)
       lName = lName + " " + Mid(Rs(I).Name, 8, 2) + " de " & Mid(Name_Month(Mid(Rs(I).Name, 5, 2)), 1, 3)
       xlSheet.Cells(3, I + 1).Value = lName
       xlSheet.Range(xlSheet.Cells(3, I + 1), xlSheet.Cells(3, I + 2)).Merge
       xlSheet.Cells(4, I + 1).Value = "ENTRADA"
    Else
       xlSheet.Cells(4, I + 1).Value = "MIN.TARD."
    End If
Next
xlSheet.Cells(4, I + 1).Value = "TOTAL MIN."
Dim lTotMin As Integer
nFil = 5
Dim lCount As Integer
lCount = 1
Do While Not Rs.EOF
   If Mid(Rs!Codigo, 1, 2) <> "**" Then
      lTotMin = 0
      For I = 0 To Rs.Fields.count - 2
         xlSheet.Cells(nFil, 1).Value = lCount
         If I < 2 Or I Mod 2 = 0 Then
            If I > 1 Then
               xlSheet.Cells(nFil, I + 2).Value = Mid(Rs(I), 1, 5)
               If Trim(Rs!entrada & "") = "" Then
                  xlSheet.Cells(nFil, I + 3).Value = "*****"
               Else
                  If Trim(Mid(Rs(I), 1, 5) & "") = "" Then
                  Else
                     If DateDiff("n", Trim(Rs!entrada & ""), Mid(Rs(I), 1, 5)) > 0 Then
                        xlSheet.Cells(nFil, I + 3).Value = DateDiff("n", Trim(Rs!entrada & ""), Mid(Rs(I), 1, 5))
                        xlSheet.Cells(nFil, I + 3).Font.Color = vbRed
                        lTotMin = lTotMin + DateDiff("n", Trim(Rs!entrada & ""), Mid(Rs(I), 1, 5))
                     Else
                        xlSheet.Cells(nFil, I + 3).Value = 0
                     End If
                  End If
               End If
            Else
               xlSheet.Cells(nFil, I + 2).Value = Trim(Rs(I) & "")
            End If
         End If
      Next
      xlSheet.Cells(nFil, I + 2).Value = lTotMin
      If lTotMin > 0 Then xlSheet.Cells(nFil, I + 2).Font.Color = vbRed
      nFil = nFil + 1
      lCount = lCount + 1
   End If
   Rs.MoveNext
Loop

xlSheet.Range("B3:B4").Merge
xlSheet.Range("C3:C4").Merge

xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(nFil - 1, Rs.Fields.count + 1)).Borders.LineStyle = xlContinuous

Rs.Close: Set Rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridlines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
Public Sub Liquidacion(lId As Double)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object


Sql$ = "Select H.*,"
Sql$ = Sql$ & "vano,vmes,vdia,gano,gmes,gdia,cano,cmes,cdia,sueldo,asigfam,promedios,promgrati,ivano,ivmes,ivdia,igano,igmes,igdia,icano,icmes,icdia,"
Sql$ = Sql$ & "rtrim(ap_pat)+' '+rtrim(ap_mat)+' '+rtrim(nom_1)+' '+rtrim(nom_2) as Nombre,"
Sql$ = Sql$ & "nro_doc,"
Sql$ = Sql$ & "(select descrip from maestros_3 where ciamaestro=P.cia+'055' and cod_maestro3=P.cargo) as Cargo,"
Sql$ = Sql$ & "fingreso,fcese,"
Sql$ = Sql$ & "(select descrip from maestros_2 where  RIGHT(ciamaestro,3)='149' and cod_maestro2=P.mot_fin_periodo) as Motivo,"
Sql$ = Sql$ & "(select flag1 from maestros_2 where ciamaestro=P.cia+'044' and cod_maestro2=P.area) as Ccosto,"
Sql$ = Sql$ & "(select descrip from maestros_2 where  RIGHT(ciamaestro,3)='069' and cod_maestro2=P.codafp) as Pension,"
Sql$ = Sql$ & "(select top 1 descrip from maestros_2 where right(ciamaestro,3)='007' and cod_maestro2=p.ctsbanco) as ctsbanco,ctsmoneda,ctsnumcta "
Sql$ = Sql$ & "from plahistorico H,plahisliquid L,planillas P "
Sql$ = Sql$ & "Where h.id_boleta = " & lId & " And h.id_boleta = L.id_boleta "
Sql$ = Sql$ & "and P.cia=H.cia and P.placod=H.placod and p.status<>'*'"

If Not (fAbrRst(Rs, Sql$)) Then
   MsgBox "Boleta de liquidación no se calculo por el sistema", vbInformation
   Rs.Close: Set Rs = Nothing
   Exit Sub
End If

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application

xlApp2.Sheets.Add
xlApp2.Sheets.Add
xlApp2.Sheets("Hoja1").Select
xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
xlApp2.Sheets("Hoja2").Select
xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)
xlApp2.Sheets("Hoja3").Select
xlApp2.Sheets("Hoja3").Move Before:=xlApp2.Sheets(3)

Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Name = "Liquidacion"

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("D:D").ColumnWidth = 8.5
xlSheet.Range("F:F").ColumnWidth = 4
xlSheet.Range("H:H").ColumnWidth = 4
xlSheet.Range("E:E").ColumnWidth = 11
xlSheet.Range("G:G").ColumnWidth = 11
xlSheet.Range("I:I").ColumnWidth = 6

xlSheet.Cells(2, 1).Value = "LIQUIDACION DE BENEFICIOS SOCIALES"
xlSheet.Cells(2, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(2, 1).Font.Bold = True
xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 10)).Merge

xlSheet.Cells(5, 1).Value = "APELLIDOS Y NOMBRES"
xlSheet.Cells(5, 4).Value = Trim(Rs!nombre & "")
xlSheet.Cells(6, 1).Value = "CODIGO TRABAJADOR"
xlSheet.Cells(6, 4).Value = "'" & Trim(Rs!PlaCod & "")
xlSheet.Cells(7, 1).Value = "DOCUMENTO DE IDENTIDAD"
xlSheet.Cells(7, 4).Value = "'" & Trim(Rs!nro_doc & "")
xlSheet.Cells(8, 1).Value = "CARGO"
xlSheet.Cells(8, 4).Value = Trim(Rs!Cargo & "")
xlSheet.Cells(9, 1).Value = "FECHA DE INGRESO"
xlSheet.Cells(9, 4).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
xlSheet.Cells(10, 1).Value = "FECHA DE CESE"
xlSheet.Cells(10, 4).Value = "'" & Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000")
xlSheet.Cells(11, 1).Value = "TIEMPO DE SERVICIO"

Dim RSL As ADODB.Recordset
Sql$ = "usp_Calcula_TSER " & Year(Rs!fIngreso) & "," & Month(Rs!fIngreso) & "," & Day(Rs!fIngreso) & "," & Year(Rs!fcese) & "," & Month(Rs!fcese) & "," & Day(Rs!fcese) & ""
If (fAbrRst(RSL, Sql$)) Then
   xlSheet.Cells(11, 4).Value = RSL!ano
   xlSheet.Cells(11, 6).Value = RSL!Mes
   xlSheet.Cells(11, 8).Value = RSL!Dia
End If
RSL.Close: Set RSL = Nothing


xlSheet.Cells(11, 5).Value = "años"
xlSheet.Cells(11, 7).Value = "meses"
xlSheet.Cells(11, 9).Value = "dias"

xlSheet.Cells(12, 1).Value = "MOTIVO DE CESE"
xlSheet.Cells(12, 4).Value = Trim(Rs!motivo & "")
xlSheet.Cells(13, 1).Value = "CENTRO DE COSTO"
xlSheet.Cells(13, 4).Value = "'" & Rs!ccosto

xlSheet.Range(xlSheet.Cells(14, 1), xlSheet.Cells(14, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(14, 1), xlSheet.Cells(14, 10)).Borders(xlEdgeBottom).Weight = xlMedium

xlSheet.Cells(16, 2).Value = "Conceptos"
xlSheet.Cells(16, 2).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(16, 2).Font.Bold = True

xlSheet.Cells(16, 5).Value = "Indemnizable"
xlSheet.Cells(16, 5).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(16, 5).Font.Bold = True
xlSheet.Cells(16, 7).Value = "Vacacional"
xlSheet.Cells(16, 7).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(16, 7).Font.Bold = True

If Rs!TipoTrab = "01" Then
   xlSheet.Cells(17, 2).Value = "Sueldo"
   xlSheet.Cells(17, 5).Value = Round(Rs!sueldo * 30, 2)
   xlSheet.Cells(17, 7).Value = Round(Rs!sueldo * 30, 2)
   
   xlSheet.Cells(18, 5).Value = Round(Rs!asigfam * 30, 2)
   xlSheet.Cells(18, 7).Value = Round(Rs!asigfam * 30, 2)
   xlSheet.Cells(19, 5).Value = Round(Rs!promedios * 30, 2)
   xlSheet.Cells(19, 7).Value = Round(Rs!promedios * 30, 2)
   xlSheet.Cells(20, 5).Value = Round(Rs!promgrati * 30, 2)
   xlSheet.Cells(21, 5).Value = Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2) + Round(Rs!promgrati * 30, 2)
   xlSheet.Cells(21, 7).Value = Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2)
Else
   xlSheet.Cells(17, 2).Value = "Jornal"
   xlSheet.Cells(17, 5).Value = Rs!sueldo
   xlSheet.Cells(17, 7).Value = Rs!sueldo
   xlSheet.Cells(18, 5).Value = Rs!asigfam
   xlSheet.Cells(18, 7).Value = Rs!asigfam
   xlSheet.Cells(19, 5).Value = Rs!promedios
   xlSheet.Cells(19, 7).Value = Rs!promedios
   xlSheet.Cells(20, 5).Value = Rs!promgrati
   xlSheet.Cells(21, 5).Value = Rs!sueldo + Rs!asigfam + Rs!promedios + Rs!promgrati
   xlSheet.Cells(21, 7).Value = Rs!sueldo + Rs!asigfam + Rs!promedios
End If
xlSheet.Cells(18, 2).Value = "Asig. Familiar"
xlSheet.Cells(19, 2).Value = "H. Extras - Bonificaciones"
xlSheet.Cells(20, 2).Value = "Prom. Gratificación"



xlSheet.Cells(21, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Cells(21, 5).Borders(xlEdgeBottom).LineStyle = xlDouble

xlSheet.Cells(21, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Cells(21, 7).Borders(xlEdgeBottom).LineStyle = xlDouble

xlSheet.Range(xlSheet.Cells(17, 5), xlSheet.Cells(21, 7)).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

Dim xLin As Integer
xLin = 23
xlSheet.Cells(xLin, 1).Value = "1.- Periodo por liquidar  ( CTS )"
xlSheet.Cells(xLin, 1).Font.Bold = True

If Rs!cano > 0 Then
   xlSheet.Cells(xLin, 6).Value = Rs!cano
   xlSheet.Cells(xLin, 7).Value = "años"
   xlSheet.Cells(xLin, 8).Value = Rs!cmes
   xlSheet.Cells(xLin, 9).Value = "meses"
   xlSheet.Cells(xLin, 10).Value = Rs!cdia
   xlSheet.Cells(xLin, 11).Value = "días"
   xlSheet.Range(xlSheet.Cells(xLin, 6), xlSheet.Cells(xLin, 11)).Font.Bold = True
Else
   xlSheet.Cells(xLin, 6).Value = Rs!cmes
   xlSheet.Cells(xLin, 7).Value = "meses"
   xlSheet.Cells(xLin, 8).Value = Rs!cdia
   xlSheet.Cells(xLin, 9).Value = "días"
   xlSheet.Range(xlSheet.Cells(xLin, 6), xlSheet.Cells(xLin, 9)).Font.Bold = True
End If

xlSheet.Cells(xLin, 10).Value = Rs!i45
xlSheet.Cells(xLin, 10).Font.Bold = True
xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "Periodo del"
If Month(Rs!fcese) = 11 Or Month(Rs!fcese) = 5 Then
   Sql$ = "select placod from plaSinPagoCTS where cia='" & wcia & "' and year(fechaproceso)= " & Year(Rs!fcese) & " and month(fechaproceso)= " & Month(Rs!fcese) - 1 & "  and placod='" & Rs!PlaCod & "' and status<>'*'"
   If (fAbrRst(RSL, Sql$)) Then
      If Month(Rs!fcese) = 5 Then
         If (Year(Rs!fcese) - 1 > Year(Rs!fIngreso)) Or (Year(Rs!fcese) = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 11) Or (Year(Rs!fcese) - 1 = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 11) Then
            xlSheet.Cells(xLin, 3).Value = "'01/11/" & Format(Year(Rs!fcese - 1), "0000")
         Else
            xlSheet.Cells(xLin, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
         End If
      Else
         If (Year(Rs!fcese) > Year(Rs!fIngreso)) Or (Year(Rs!fcese) = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 5) Then
            xlSheet.Cells(xLin, 3).Value = "'01/05/" & Format(Year(Rs!fcese), "0000")
         Else
            xlSheet.Cells(xLin, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
         End If
      End If
   Else
      If Month(Rs!fcese) = 5 Then
         xlSheet.Cells(xLin, 3).Value = "'01/05/" & Format(Year(Rs!fcese), "0000")
      Else
         xlSheet.Cells(xLin, 3).Value = "'01/11/" & Format(Year(Rs!fcese), "0000")
      End If
   End If
   RSL.Close: Set RSL = Nothing
Else
   If Month(Rs!fcese) < 5 Then
      If (Year(Rs!fcese) - 1 > Year(Rs!fIngreso)) Or (Year(Rs!fcese) = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 11) Or (Year(Rs!fcese) - 1 = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 11) Then
         xlSheet.Cells(xLin, 3).Value = "'01/11/" & Format(Year(Rs!fcese) - 1, "0000")
      Else
         xlSheet.Cells(xLin, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
      End If
   Else
      If (Year(Rs!fcese) > Year(Rs!fIngreso)) Or (Year(Rs!fcese) = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 5) Then
         xlSheet.Cells(xLin, 3).Value = "'01/05/" & Format(Year(Rs!fcese), "0000")
      Else
         xlSheet.Cells(xLin, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
      End If
   End If
End If

xlSheet.Cells(xLin, 4).Value = "al"
xlSheet.Cells(xLin, 5).Value = "'" & Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000")

xLin = xLin + 2
If Rs!icano <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( " & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2) + Round(Rs!promgrati * 30, 2))) & " * " & Trim(Str(Rs!cano)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( " & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios + Rs!promgrati, 2))) & " * 30 * " & Trim(Str(Rs!cano)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!icano
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If

If Rs!icmes <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2) + Round(Rs!promgrati * 30, 2))) & ") / 12) * " & Trim(Str(Rs!cmes)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios + Rs!promgrati, 2))) & " * 30 ) / 12) * " & Trim(Str(Rs!cmes)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!icmes
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If

If Rs!icdia <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2) + Round(Rs!promgrati * 30, 2))) & ") / 360) * " & Trim(Str(Rs!cdia)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios + Rs!promgrati, 2))) & " * 30 )/360) * " & Trim(Str(Rs!cdia)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!icdia
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If
xLin = xLin + 1

xlSheet.Cells(xLin, 1).Value = "2.- Vacaciones Truncas"
xlSheet.Cells(xLin, 1).Font.Bold = True

xlSheet.Cells(xLin, 10).Value = Rs!i39
xlSheet.Cells(xLin, 10).Font.Bold = True
xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

If Rs!Vano > 0 Then
   xlSheet.Cells(xLin, 6).Value = Rs!Vano
   xlSheet.Cells(xLin, 7).Value = "años"
   xlSheet.Cells(xLin, 8).Value = Rs!Vmes
   xlSheet.Cells(xLin, 9).Value = "meses"
   xlSheet.Cells(xLin, 10).Value = Rs!Vdia
   xlSheet.Cells(xLin, 11).Value = "días"
   xlSheet.Range(xlSheet.Cells(xLin, 6), xlSheet.Cells(xLin, 11)).Font.Bold = True
Else
   xlSheet.Cells(xLin, 6).Value = Rs!Vmes
   xlSheet.Cells(xLin, 7).Value = "meses"
   xlSheet.Cells(xLin, 8).Value = Rs!Vdia
   xlSheet.Cells(xLin, 9).Value = "días"
   xlSheet.Range(xlSheet.Cells(xLin, 6), xlSheet.Cells(xLin, 9)).Font.Bold = True
End If

xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "Periodo del"

Dim lFechaVac As Date


If Year(Rs!fcese) = Year(Rs!fIngreso) Then
   lFechaVac = Rs!fIngreso
ElseIf DateDiff("yyyy", Rs!fIngreso, Rs!fcese) < 1 Then
   lFechaVac = Rs!fIngreso
ElseIf DateDiff("m", Rs!fIngreso, Rs!fcese) < 12 Then
   lFechaVac = Rs!fIngreso
Else
   lFechaVac = Rs!fcese
   If Rs!Vano <> 0 Then lFechaVac = DateAdd("y", -Rs!Vano, lFechaVac)
   If Rs!Vmes <> 0 Then lFechaVac = DateAdd("m", -Rs!Vmes, lFechaVac)
   If Rs!Vdia <> 0 Then lFechaVac = DateAdd("d", -(Rs!Vdia - 1), lFechaVac)
End If

xlSheet.Cells(xLin, 3).Value = lFechaVac

xlSheet.Cells(xLin, 5).Value = "al"
xlSheet.Cells(xLin, 5).Value = "'" & Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000")

xLin = xLin + 2
If Rs!iVano <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( " & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2))) & " * " & Trim(Str(Rs!Vano)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( " & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios, 2))) & " * 30 * " & Trim(Str(Rs!Vano)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!iVano
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If
If Rs!iVmes <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2))) & ") / 12) * " & Trim(Str(Rs!Vmes)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios, 2))) & " * 30 )/12) * " & Trim(Str(Rs!Vmes)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!iVmes
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If
If Rs!iVdia <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2))) & ") / 360) * " & Trim(Str(Rs!Vdia)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios, 2))) & " * 30 )/360) * " & Trim(Str(Rs!Vdia)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!iVdia
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If
xLin = xLin + 1

xlSheet.Cells(xLin, 1).Value = "3.- Gratificaciones Truncas"
xlSheet.Cells(xLin, 1).Font.Bold = True

xlSheet.Cells(xLin, 10).Value = Rs!i40
xlSheet.Cells(xLin, 10).Font.Bold = True
xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

If Rs!gano > 0 Then
   xlSheet.Cells(xLin, 6).Value = Rs!gano
   xlSheet.Cells(xLin, 7).Value = "años"
   xlSheet.Cells(xLin, 8).Value = Rs!gmes
   xlSheet.Cells(xLin, 9).Value = "meses"
   xlSheet.Cells(xLin, 10).Value = Rs!gdia
   xlSheet.Cells(xLin, 11).Value = "días"
   xlSheet.Range(xlSheet.Cells(xLin, 6), xlSheet.Cells(xLin, 11)).Font.Bold = True
Else
   xlSheet.Cells(xLin, 6).Value = Rs!gmes
   xlSheet.Cells(xLin, 7).Value = "meses"
   xlSheet.Cells(xLin, 8).Value = Rs!gdia
   xlSheet.Cells(xLin, 9).Value = "días"
   xlSheet.Range(xlSheet.Cells(xLin, 6), xlSheet.Cells(xLin, 9)).Font.Bold = True
End If

xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "Periodo del"

If Month(Rs!fcese) <= 7 Then
   If Year(Rs!fcese) > Year(Rs!fIngreso) Then
      xlSheet.Cells(xLin, 3).Value = "'01/01/" & Format(Year(Rs!fcese), "0000")
   Else
      xlSheet.Cells(xLin, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
   End If
Else
   If (Year(Rs!fcese) > Year(Rs!fIngreso)) Or (Year(Rs!fcese) = Year(Rs!fIngreso) And Month(Rs!fIngreso) < 7) Then
      xlSheet.Cells(xLin, 3).Value = "'01/07/" & Format(Year(Rs!fcese), "0000")
   Else
      xlSheet.Cells(xLin, 3).Value = "'" & Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000")
   End If
End If
   
xlSheet.Cells(xLin, 4).Value = "al"
xlSheet.Cells(xLin, 5).Value = "'" & Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000")

xLin = xLin + 2
If Rs!igmes <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2))) & ") / 6) * " & Trim(Str(Rs!gmes)) & " )"
   Else
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios, 2))) & " * 30 ) / 6) * " & Trim(Str(Rs!gmes)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!igmes
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If
If Rs!igdia <> 0 Then
   If Rs!TipoTrab = "01" Then
      xlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo * 30, 2) + Round(Rs!asigfam * 30, 2) + Round(Rs!promedios * 30, 2))) & ") / 180) * " & Trim(Str(Rs!gdia)) & " )"
   Else
      BxlSheet.Cells(xLin, 2).Value = "( ((" & Trim(Str(Round(Rs!sueldo + Rs!asigfam + Rs!promedios, 2))) & " * 30 ) / 180) * " & Trim(Str(Rs!gdia)) & " )"
   End If
   xlSheet.Cells(xLin, 4).Value = "'="
   xlSheet.Cells(xLin, 5).Value = Rs!igdia
   xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
End If
xLin = xLin + 1

Dim lItem As Integer
Dim I As Integer
lItem = 4

If Rs!i40 <> 0 And Rs!i30 <> 0 Then
   xlSheet.Cells(xLin, 1).Value = Trim(Str(lItem)) & ".- Bonificación Extraordinaria"
   xlSheet.Cells(xLin, 1).Font.Bold = True

   xlSheet.Cells(xLin, 10).Value = Rs!i30
   xlSheet.Cells(xLin, 10).Font.Bold = True
   xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
   xLin = xLin + 1
   xlSheet.Cells(xLin, 2).Value = "Ley 29351"
   xlSheet.Cells(xLin, 5).Value = Rs!i30
   xLin = xLin + 2
   lItem = lItem + 1
End If

Dim lBolNol As Double
lBolNol = Rs!i01 + Rs!i02 + Rs!i03 + Rs!i04 + Rs!i05 + Rs!i06 + Rs!i07 + Rs!i08 + Rs!i09 + Rs!i10 + Rs!i11 + Rs!i12 + Rs!i13 + Rs!i14 + Rs!i15 + Rs!I16 + Rs!i17 + Rs!i18 + Rs!i19 + Rs!i20 + Rs!i21 + Rs!i22 + Rs!i23
lBolNol = lBolNol + Rs!I24 + Rs!i25 + Rs!i26 + Rs!i27 + Rs!i28 + Rs!i29 + Rs!i31 + Rs!i32 + Rs!i33 + Rs!i34 + Rs!i35 + Rs!i36 + Rs!i37 + Rs!i38 + Rs!i41 + Rs!i42 + Rs!i43 + Rs!i44 + Rs!i46 + Rs!i47 + Rs!i48 + Rs!i49 + Rs!i50
If lBolNol Then
   xlSheet.Cells(xLin, 1).Value = Trim(Str(lItem)) & ".- Boleta Normal"
   xlSheet.Cells(xLin, 1).Font.Bold = True
   
   xlSheet.Cells(xLin, 10).Value = Rs!totaling - Rs!i30 - Rs!i39 - Rs!i40 - Rs!i45
   xlSheet.Cells(xLin, 10).Font.Bold = True
   xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

   xLin = xLin + 1
   Dim lField As String
   For I = 1 To 50
       If I = 30 Or I = 39 Or I = 40 Or I = 45 Then
       Else
          lField = "i" & Format(I, "00")
          If Rs(lField) <> 0 Then
             Sql$ = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno='" & Format(I, "00") & "' and status<>'*'"
             If (fAbrRst(RSL, Sql$)) Then xlSheet.Cells(xLin, 2).Value = Trim(RSL(0) & "")
             RSL.Close: Set RSL = Nothing
             Select Case I
                Case Is = 1: xlSheet.Cells(xLin, 4).Value = Rs!h01
                Case Is = 9: xlSheet.Cells(xLin, 4).Value = Rs!h02
                Case Is = 10: xlSheet.Cells(xLin, 4).Value = Rs!h10
                Case Is = 11: xlSheet.Cells(xLin, 4).Value = Rs!h11
                Case Is = 12: xlSheet.Cells(xLin, 4).Value = Rs!h03
                Case Is = 21: xlSheet.Cells(xLin, 4).Value = Rs!h17
                Case Is = 24: xlSheet.Cells(xLin, 4).Value = Rs!h18
                Case Is = 25: xlSheet.Cells(xLin, 4).Value = Rs!h19
                Case Is = 32: xlSheet.Cells(xLin, 4).Value = Rs!h05
                Case Is = 38: xlSheet.Cells(xLin, 4).Value = Rs!h24
                Case Is = 42: xlSheet.Cells(xLin, 4).Value = Rs!h26
                Case Is = 43: xlSheet.Cells(xLin, 4).Value = Rs!h27
             End Select
             xlSheet.Cells(xLin, 4).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
             xlSheet.Cells(xLin, 5).Value = Rs(lField)
             xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
             xLin = xLin + 1
          End If
       End If
   Next
End If

xLin = xLin + 1
xlSheet.Cells(xLin, 1).Value = "TOTAL BRUTO"
xlSheet.Cells(xLin, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(xLin, 1).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 9)).Merge
xlSheet.Cells(xLin, 10).Value = Rs!totaling
xlSheet.Cells(xLin, 10).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).Borders.LineStyle = xlContinuous

xLin = xLin + 2
xlSheet.Cells(xLin, 1).Value = "DEDUCCIONES"
xlSheet.Cells(xLin, 1).Font.Bold = True
xlSheet.Cells(xLin, 10).Value = Rs!totalded * -1
xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red](#,##0.00) "
xlSheet.Cells(xLin, 10).Font.Bold = True
xLin = xLin + 1
For I = 1 To 20
    lField = "d" & Format(I, "00")
    If Rs(lField) <> 0 Then
       If I = 4 Or I = 11 Then
          xlSheet.Cells(xLin, 2).Value = Trim(Rs!PENSION & "")
       Else
          Sql$ = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & Format(I, "00") & "' and status<>'*'"
          If (fAbrRst(RSL, Sql$)) Then xlSheet.Cells(xLin, 2).Value = Trim(RSL(0) & "")
          RSL.Close: Set RSL = Nothing
       End If
       xlSheet.Cells(xLin, 5).Value = Rs(lField)
       xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
       xLin = xLin + 1
    End If
Next
xLin = xLin + 1

xlSheet.Cells(xLin, 1).Value = "NETO A PAGAR"
xlSheet.Cells(xLin, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(xLin, 1).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 9)).Merge
xlSheet.Cells(xLin, 10).Value = Rs!totaling - Rs!totalded
xlSheet.Cells(xLin, 10).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).Borders.LineStyle = xlContinuous

xLin = xLin + 2
xlSheet.Cells(xLin, 1).Value = "APORTES"
xlSheet.Cells(xLin, 1).Font.Bold = True
xlSheet.Cells(xLin, 10).NumberFormat = "#,##0.00_ ;[Red](#,##0.00) "
xlSheet.Cells(xLin, 10).Font.Bold = True
xLin = xLin + 1
For I = 1 To 20
    lField = "a" & Format(I, "00")
    If Rs(lField) <> 0 Then
       Sql$ = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & Format(I, "00") & "' and status<>'*'"
       If (fAbrRst(RSL, Sql$)) Then xlSheet.Cells(xLin, 2).Value = Trim(RSL(0) & "")
       RSL.Close: Set RSL = Nothing
       xlSheet.Cells(xLin, 5).Value = Rs(lField)
       xlSheet.Cells(xLin, 5).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
       xLin = xLin + 1
    End If
Next
xLin = xLin + 1
xlSheet.Cells(xLin, 1).Value = "TOTAL APORTES"
xlSheet.Cells(xLin, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(xLin, 1).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 9)).Merge
xlSheet.Cells(xLin, 10).Value = Rs!totalapo
xlSheet.Cells(xLin, 10).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).Borders.LineStyle = xlContinuous


xLin = xLin + 2
Dim montolet As String
If Rs!totaling - Rs!totalded <> 0 Then montolet = monto_palabras(Rs!totaling - Rs!totalded) Else montolet = ""
Sql$ = "He recibido la cantidad de S/." & Format(Rs!totaling - Rs!totalded, "###,###,###.00") & " (" & montolet & " Soles) "
Sql$ = Sql$ & "correspondiente a mis beneficios sociales conforme a las leyes vigentes, no teniendo nada que reclamar a "
Sql$ = Sql$ & NOMBREEMPRESA & " por este concepto."
xlSheet.Cells(xLin, 1).Value = Sql$
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 1)).RowHeight = 50
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).Merge

xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).VerticalAlignment = xlJustify
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).HorizontalAlignment = xlJustify
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 10)).Font.Size = 9

xLin = xLin + 2
xlSheet.Cells(xLin, 1).NumberFormat = "[$-280A]d"" de ""mmmm"" de ""yyyy;@"
xlSheet.Cells(xLin, 1).Value = Rs!fcese
'xlSheet.Cells(xLin, 1).Value = Format(Day(rs!fcese), "00") & "/" & Format(Month(rs!fcese), "00") & "/" & Format(Year(rs!fcese), "0000")
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 3)).Merge
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 3)).HorizontalAlignment = xlCenter

xLin = xLin + 3
xlSheet.Cells(xLin, 5).Value = Trim(Rs!nombre & "")
xlSheet.Range(xlSheet.Cells(xLin, 5), xlSheet.Cells(xLin, 10)).Merge
xlSheet.Range(xlSheet.Cells(xLin, 5), xlSheet.Cells(xLin, 5)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(xLin, 5), xlSheet.Cells(xLin, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
xLin = xLin + 1
xlSheet.Cells(xLin, 5).Value = "D.N.I   " & Trim(Rs!nro_doc & "")
xlSheet.Range(xlSheet.Cells(xLin, 5), xlSheet.Cells(xLin, 10)).Merge
xlSheet.Range(xlSheet.Cells(xLin, 5), xlSheet.Cells(xLin, 5)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(xLin - 1, 5), xlSheet.Cells(xLin, 5)).Font.Bold = True

Call Certificado_Trabajo(Trim(Rs!nombre & ""), Trim(Rs!Cargo & ""), Format(Day(Rs!fIngreso), "00") & "/" & Format(Month(Rs!fIngreso), "00") & "/" & Format(Year(Rs!fIngreso), "0000"), Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000"))
Call Constancia_CTS(Trim(Rs!nombre & ""), Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000"), Trim(Rs!ctsbanco & ""), Trim(Rs!ctsmoneda & ""), Trim(Rs!ctsnumcta & ""), Trim(Rs!nro_doc & ""))


For I = 1 To 3
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridlines = False
   'xlApp2.ActiveWindow.Zoom = 80
Next
xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
Private Sub Certificado_Trabajo(nombre As String, Cargo As String, fi As Date, fc As Date)
Dim xLin As Integer
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "Certificado"
xlSheet.Range("A:A").ColumnWidth = 5
xlSheet.Range("B:B").ColumnWidth = 2
xlSheet.Range("C:C").ColumnWidth = 4
xlSheet.Range("D:I").ColumnWidth = 14

xLin = 15
xlSheet.Cells(xLin, 1).Value = "CERTIFICADO"
xlSheet.Cells(xLin, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(xLin, 1).Font.Size = 24
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 9)).Merge
xlSheet.Range(xlSheet.Cells(xLin, 1), xlSheet.Cells(xLin, 1)).RowHeight = 29
xlSheet.Cells(xLin, 1).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(xLin, 1).Font.Bold = True
xLin = xLin + 3
xlSheet.Cells(xLin, 3).Value = "Por el presente documento certificamos que él (la)  señor (ita):"
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 2
xlSheet.Cells(xLin, 3).Value = nombre
xlSheet.Cells(xLin, 3).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(xLin, 3).HorizontalAlignment = xlCenter
xlSheet.Cells(xLin, 3).Font.Bold = True
xlSheet.Cells(xLin, 3).Font.Size = 14
xlSheet.Cells(xLin, 3).Font.Bold = True
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 3
xlSheet.Cells(xLin, 3).Value = "Ha trabajado en nuestra empresa, desempeñando el  cargo  de"
xlSheet.Cells(xLin, 3).HorizontalAlignment = xlDistributed
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 3)).RowHeight = 15
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 1
xlSheet.Cells(xLin, 3).Value = Cargo & " desde el " & fi & " al " & fc
xlSheet.Cells(xLin, 3).HorizontalAlignment = xlDistributed
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 3)).RowHeight = 15
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Cells(xLin, 3).Characters(Start:=1, Length:=Len(Trim(Cargo))).Font.Bold = True
xlSheet.Cells(xLin, 3).Characters(Start:=11 + Len(Trim(Cargo)), Length:=10).Font.Bold = True
xlSheet.Cells(xLin, 3).Characters(Start:=25 + Len(Trim(Cargo)), Length:=10).Font.Bold = True
xLin = xLin + 2
xlSheet.Cells(xLin, 3).Value = "Durante el tiempo de permanencia en nuestra empresa ha demostrado"
xlSheet.Cells(xLin, 3).HorizontalAlignment = xlDistributed
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 3)).RowHeight = 15
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 1
xlSheet.Cells(xLin, 3).Value = "eficiencia, honestidad y responsabilidad en el desempeño de sus labores."
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 2
xlSheet.Cells(xLin, 3).Value = "Se expide el presente certificado a solicitud del interesado, para los fines"
xlSheet.Cells(xLin, 3).HorizontalAlignment = xlDistributed
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 3)).RowHeight = 15
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 1
xlSheet.Cells(xLin, 3).Value = "que estime conveniente."
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 9)).Merge
xLin = xLin + 3
xlSheet.Cells(xLin, 6).Value = "Lima,"
xlSheet.Cells(xLin, 6).Font.Size = 12
xlSheet.Cells(xLin, 6).HorizontalAlignment = xlRight
xlSheet.Cells(xLin, 7).NumberFormat = "[$-280A]d"" de ""mmmm"" de ""yyyy;@"
xlSheet.Cells(xLin, 7).Value = fc
xlSheet.Cells(xLin, 7).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 7), xlSheet.Cells(xLin, 9)).Merge
xlSheet.Cells(xLin, 7).HorizontalAlignment = xlLeft
End Sub
Private Sub Constancia_CTS(nombre As String, fc As Date, ctsbanco As String, ctsmoneda As String, ctsnumcta As String, DNI As String)
Dim xLin As Integer
Set xlSheet = xlApp2.Worksheets("HOJA3")
xlSheet.Name = "Constancia"

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 15
xlSheet.Range("D:D").ColumnWidth = 3

xLin = 13

xlSheet.Cells(xLin, 2).Value = "Lima,"
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).HorizontalAlignment = xlRight
xlSheet.Cells(xLin, 3).NumberFormat = "[$-280A]d"" de ""mmmm"" de ""yyyy;@"
xlSheet.Cells(xLin, 3).Value = fc
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Range(xlSheet.Cells(xLin, 3), xlSheet.Cells(xLin, 5)).Merge
xlSheet.Cells(xLin, 3).HorizontalAlignment = xlLeft

xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "Señores"
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = ctsbanco
xlSheet.Cells(xLin, 2).Font.Bold = True
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "Ciudad.-"
xlSheet.Cells(xLin, 2).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "A/A: Dpto. de C.T.S."
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "REF : RETIRO DE COMPENSACION POR TIEMPO DE SERVICIOS"
xlSheet.Cells(xLin, 2).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).Font.Bold = True
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "C.T.S."
xlSheet.Cells(xLin, 2).Font.Underline = xlUnderlineStyleSingle
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).Font.Bold = True
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "De nuestra consideración:"
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "Mediante el presente comunicamos a ustedes que la persona que a"
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "continuación se detalla ha dejado de laborar en nuestra empresa:"
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "APELLIDOS  Y  NOMBRES"
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).Font.Bold = True
xlSheet.Cells(xLin, 4).Value = ":"
xlSheet.Cells(xLin, 4).Font.Size = 12
xlSheet.Cells(xLin, 4).Font.Bold = True
xlSheet.Cells(xLin, 5).Value = nombre
xlSheet.Cells(xLin, 5).Font.Size = 12
xlSheet.Cells(xLin, 5).Font.Bold = True
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "DNI"
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).Font.Bold = True
xlSheet.Cells(xLin, 4).Value = ":"
xlSheet.Cells(xLin, 4).Font.Size = 12
xlSheet.Cells(xLin, 4).Font.Bold = True
xlSheet.Cells(xLin, 5).Value = DNI
xlSheet.Cells(xLin, 5).Font.Size = 12
xlSheet.Cells(xLin, 5).Font.Bold = True
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "CTA C.T.S."
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).Font.Bold = True
If ctsmoneda = "S/." Then
   xlSheet.Cells(xLin, 3).Value = "M.N. (" & ctsmoneda & ")"
Else
   xlSheet.Cells(xLin, 3).Value = "M.E. (" & ctsmoneda & ")"
End If
xlSheet.Cells(xLin, 3).Font.Size = 12
xlSheet.Cells(xLin, 3).Font.Bold = True
xlSheet.Cells(xLin, 4).Value = ":"
xlSheet.Cells(xLin, 4).Font.Size = 12
xlSheet.Cells(xLin, 4).Font.Bold = True
xlSheet.Cells(xLin, 5).Value = "'" & ctsnumcta
xlSheet.Cells(xLin, 5).Font.Size = 12
xlSheet.Cells(xLin, 5).Font.Bold = True
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "FECHA DE CESE"
xlSheet.Cells(xLin, 2).Font.Size = 12
xlSheet.Cells(xLin, 2).Font.Bold = True
xlSheet.Cells(xLin, 4).Value = ":"
xlSheet.Cells(xLin, 4).Font.Size = 12
xlSheet.Cells(xLin, 4).Font.Bold = True
xlSheet.Cells(xLin, 5).Value = fc
xlSheet.Cells(xLin, 5).Font.Size = 12
xlSheet.Cells(xLin, 5).Font.Bold = True
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "Por este motivo, agradeceremos se sirvan brindarle las facilidades de"
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 1
xlSheet.Cells(xLin, 2).Value = "caso para el retiro de su fondo  de compensación."
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "Agradeciéndoles la atención al presente, quedamos de Uds."
xlSheet.Cells(xLin, 2).Font.Size = 12
xLin = xLin + 2
xlSheet.Cells(xLin, 2).Value = "Atentamente,"
xlSheet.Cells(xLin, 2).Font.Size = 12
End Sub

Public Function Letra_ColumnaExcel(ByVal Columna As Integer) As String
        Select Case Columna
            Case 1
                Letra_ColumnaExcel = "A"
            Case 2
                Letra_ColumnaExcel = "B"
            Case 3
                Letra_ColumnaExcel = "C"
            Case 4
                Letra_ColumnaExcel = "D"
            Case 5
                Letra_ColumnaExcel = "E"
            Case 6
                Letra_ColumnaExcel = "F"
            Case 7
                Letra_ColumnaExcel = "G"
            Case 8
                Letra_ColumnaExcel = "H"
            Case 9
                Letra_ColumnaExcel = "I"
            Case 10
                Letra_ColumnaExcel = "J"
            Case 11
                Letra_ColumnaExcel = "K"
            Case 12
                Letra_ColumnaExcel = "L"
            Case 13
                Letra_ColumnaExcel = "M"
            Case 14
                Letra_ColumnaExcel = "N"
            Case 15
                Letra_ColumnaExcel = "O"
            Case 16
                Letra_ColumnaExcel = "P"
            Case 17
                Letra_ColumnaExcel = "Q"
            Case 18
                Letra_ColumnaExcel = "R"
            Case 19
                Letra_ColumnaExcel = "S"
            Case 20
                Letra_ColumnaExcel = "T"
            Case 21
                Letra_ColumnaExcel = "U"
            Case 22
                Letra_ColumnaExcel = "V"
            Case 23
                Letra_ColumnaExcel = "W"
            Case 24
                Letra_ColumnaExcel = "X"
            Case 25
                Letra_ColumnaExcel = "Y"
            Case 26
                Letra_ColumnaExcel = "Z"
            Case 27
                Letra_ColumnaExcel = "AA"
            Case 28
                Letra_ColumnaExcel = "AB"
            Case 29
                Letra_ColumnaExcel = "AC"
            Case 30
                Letra_ColumnaExcel = "AD"
            Case 31
                Letra_ColumnaExcel = "AE"
            Case 32
                Letra_ColumnaExcel = "AF"
            Case 33
                Letra_ColumnaExcel = "AG"
            Case 34
                Letra_ColumnaExcel = "AH"
            Case 35
                Letra_ColumnaExcel = "AI"
            Case 36
                Letra_ColumnaExcel = "AJ"
            Case 37
                Letra_ColumnaExcel = "AK"
            Case 38
                Letra_ColumnaExcel = "AL"
            Case 39
                Letra_ColumnaExcel = "AM"
            Case 40
                Letra_ColumnaExcel = "AN"
            Case 41
                Letra_ColumnaExcel = "AO"
            Case 42
                Letra_ColumnaExcel = "AP"
            Case 43
                Letra_ColumnaExcel = "AQ"
            Case 44
                Letra_ColumnaExcel = "AR"
            Case 45
                Letra_ColumnaExcel = "AS"
            Case 46
                Letra_ColumnaExcel = "AT"
            Case 47
                Letra_ColumnaExcel = "AU"
            Case 48
                Letra_ColumnaExcel = "AV"
            Case 49
                Letra_ColumnaExcel = "AW"
            Case 50
                Letra_ColumnaExcel = "AX"
            Case 51
                Letra_ColumnaExcel = "AY"
            Case 52
                Letra_ColumnaExcel = "AZ"
            Case 53
                Letra_ColumnaExcel = "BA"
            Case 54
                Letra_ColumnaExcel = "BB"
            Case 55
                Letra_ColumnaExcel = "BC"
            Case 56
                Letra_ColumnaExcel = "BD"
            Case 57
                Letra_ColumnaExcel = "BE"
            Case 58
                Letra_ColumnaExcel = "BF"
            Case 59
                Letra_ColumnaExcel = "BG"
            Case 60
                Letra_ColumnaExcel = "BH"
            Case 61
                Letra_ColumnaExcel = "BI"
            Case 62
                Letra_ColumnaExcel = "BJ"
            Case 63
                Letra_ColumnaExcel = "BK"
            Case 64
                Letra_ColumnaExcel = "BL"
            Case 65
                Letra_ColumnaExcel = "BM"
            Case 66
                Letra_ColumnaExcel = "BN"
            Case 67
                Letra_ColumnaExcel = "BO"
            Case 68
                Letra_ColumnaExcel = "BP"
            Case 69
                Letra_ColumnaExcel = "BQ"
            Case 70
                Letra_ColumnaExcel = "BR"
            Case 71
                Letra_ColumnaExcel = "BS"
            Case 72
                Letra_ColumnaExcel = "BT"
            Case 73
                Letra_ColumnaExcel = "BU"
            Case 74
                Letra_ColumnaExcel = "BV"
            Case 75
                Letra_ColumnaExcel = "BW"
            Case 76
                Letra_ColumnaExcel = "BX"
            Case 77
                Letra_ColumnaExcel = "BY"
            Case 78
                Letra_ColumnaExcel = "BZ"
            Case 79
                Letra_ColumnaExcel = "CA"
            Case 80
                Letra_ColumnaExcel = "CB"
            Case 81
                Letra_ColumnaExcel = "CC"
            Case 82
                Letra_ColumnaExcel = "CD"
            Case 83
                Letra_ColumnaExcel = "CE"
            Case 84
                Letra_ColumnaExcel = "CF"
            Case 85
                Letra_ColumnaExcel = "CG"
            Case 86
                Letra_ColumnaExcel = "CH"
            Case 87
                Letra_ColumnaExcel = "CI"
            Case 88
                Letra_ColumnaExcel = "CJ"
            Case 89
                Letra_ColumnaExcel = "CK"
            Case 90
                Letra_ColumnaExcel = "CL"
            Case 91
                Letra_ColumnaExcel = "CM"
            Case 92
                Letra_ColumnaExcel = "CN"
            Case 93
                Letra_ColumnaExcel = "CO"
            Case 94
                Letra_ColumnaExcel = "CP"
            Case 95
                Letra_ColumnaExcel = "CQ"
            Case 96
                Letra_ColumnaExcel = "CR"
            Case 97
                Letra_ColumnaExcel = "CS"
            Case 98
                Letra_ColumnaExcel = "CT"
            Case 99
                Letra_ColumnaExcel = "CU"
            Case 100
                Letra_ColumnaExcel = "CV"
            Case 101
                Letra_ColumnaExcel = "CW"
            Case 102
                Letra_ColumnaExcel = "CX"
            Case 103
                Letra_ColumnaExcel = "CY"
            Case 104
                Letra_ColumnaExcel = "CZ"
            Case 105
                Letra_ColumnaExcel = "DA"
            Case 106
                Letra_ColumnaExcel = "DB"
            Case 107
                Letra_ColumnaExcel = "DC"
            Case 108
                Letra_ColumnaExcel = "DD"
            Case 109
                Letra_ColumnaExcel = "DE"
            Case 110
                Letra_ColumnaExcel = "DF"
            Case 111
                Letra_ColumnaExcel = "DG"
            Case 112
                Letra_ColumnaExcel = "DH"
            Case 113
                Letra_ColumnaExcel = "DI"
            Case 114
                Letra_ColumnaExcel = "DJ"
            Case 115
                Letra_ColumnaExcel = "DK"
            Case 116
                Letra_ColumnaExcel = "DL"
            Case 117
                Letra_ColumnaExcel = "DM"
            Case 118
                Letra_ColumnaExcel = "DN"
            Case 119
                Letra_ColumnaExcel = "DO"
            Case 120
                Letra_ColumnaExcel = "DP"
            Case 121
                Letra_ColumnaExcel = "DQ"
            Case 122
                Letra_ColumnaExcel = "DR"
            Case 123
                Letra_ColumnaExcel = "DS"
            Case 124
                Letra_ColumnaExcel = "DT"
            Case 125
                Letra_ColumnaExcel = "DU"
            Case 126
                Letra_ColumnaExcel = "DV"
            Case 127
                Letra_ColumnaExcel = "DW"
            Case 128
                Letra_ColumnaExcel = "DX"
            Case 129
                Letra_ColumnaExcel = "DY"
            Case 130
                Letra_ColumnaExcel = "DZ"
            Case 131
                Letra_ColumnaExcel = "EA"
            Case 132
                Letra_ColumnaExcel = "EB"
            Case 133
                Letra_ColumnaExcel = "EC"
            Case 134
                Letra_ColumnaExcel = "ED"
            Case 135
                Letra_ColumnaExcel = "EE"
            Case 136
                Letra_ColumnaExcel = "EF"
            Case 137
                Letra_ColumnaExcel = "EG"
            Case 138
                Letra_ColumnaExcel = "EH"
            Case 139
                Letra_ColumnaExcel = "EI"
            Case 140
                Letra_ColumnaExcel = "EJ"
            Case 141
                Letra_ColumnaExcel = "EK"
            Case 142
                Letra_ColumnaExcel = "EL"
            Case 143
                Letra_ColumnaExcel = "EM"
            Case 144
                Letra_ColumnaExcel = "EN"
            Case 145
                Letra_ColumnaExcel = "EO"
            Case 146
                Letra_ColumnaExcel = "EP"
            Case 147
                Letra_ColumnaExcel = "EQ"
            Case 148
                Letra_ColumnaExcel = "ER"
            Case 149
                Letra_ColumnaExcel = "ES"
            Case 150
                Letra_ColumnaExcel = "ET"
            Case 151
                Letra_ColumnaExcel = "EU"
            Case 152
                Letra_ColumnaExcel = "EV"
            Case 153
                Letra_ColumnaExcel = "EW"
            Case 154
                Letra_ColumnaExcel = "EX"
            Case 155
                Letra_ColumnaExcel = "EY"
            Case 156
                Letra_ColumnaExcel = "EZ"
            Case 157
                Letra_ColumnaExcel = "FA"
            Case 158
                Letra_ColumnaExcel = "FB"
            Case 159
                Letra_ColumnaExcel = "FC"
            Case 160
                Letra_ColumnaExcel = "FD"
            Case 161
                Letra_ColumnaExcel = "FE"
            Case 162
                Letra_ColumnaExcel = "FF"
            Case 163
                Letra_ColumnaExcel = "FG"
            Case 164
                Letra_ColumnaExcel = "FH"
            Case 165
                Letra_ColumnaExcel = "FI"
            Case 166
                Letra_ColumnaExcel = "FJ"
            Case 167
                Letra_ColumnaExcel = "FK"
            Case 168
                Letra_ColumnaExcel = "FL"
            Case 169
                Letra_ColumnaExcel = "FM"
            Case 170
                Letra_ColumnaExcel = "FN"
            Case 171
                Letra_ColumnaExcel = "FO"
            Case 172
                Letra_ColumnaExcel = "FP"
            Case 173
                Letra_ColumnaExcel = "FQ"
            Case 174
                Letra_ColumnaExcel = "FR"
            Case 175
                Letra_ColumnaExcel = "FS"
            Case 176
                Letra_ColumnaExcel = "FT"
            Case 177
                Letra_ColumnaExcel = "FU"
            Case 178
                Letra_ColumnaExcel = "FV"
            Case 179
                Letra_ColumnaExcel = "FW"
            Case 180
                Letra_ColumnaExcel = "FX"
            Case 181
                Letra_ColumnaExcel = "FY"
            Case 182
                Letra_ColumnaExcel = "FZ"
            Case 183
                Letra_ColumnaExcel = "GA"
            Case 184
                Letra_ColumnaExcel = "GB"
            Case 185
                Letra_ColumnaExcel = "GC"
            Case 186
                Letra_ColumnaExcel = "GD"
            Case 187
                Letra_ColumnaExcel = "GE"
            Case 188
                Letra_ColumnaExcel = "GF"
            Case 189
                Letra_ColumnaExcel = "GG"
            Case 190
                Letra_ColumnaExcel = "GH"
            Case 191
                Letra_ColumnaExcel = "GI"
            Case 192
                Letra_ColumnaExcel = "GJ"
            Case 193
                Letra_ColumnaExcel = "GK"
            Case 194
                Letra_ColumnaExcel = "GL"
            Case 195
                Letra_ColumnaExcel = "GM"
            Case 196
                Letra_ColumnaExcel = "GN"
            Case 197
                Letra_ColumnaExcel = "GO"
            Case 198
                Letra_ColumnaExcel = "GP"
            Case 199
                Letra_ColumnaExcel = "GQ"
            Case 200
                Letra_ColumnaExcel = "GR"
            Case 201
                Letra_ColumnaExcel = "GS"
            Case 202
                Letra_ColumnaExcel = "GT"
            Case 203
                Letra_ColumnaExcel = "GU"
            Case 204
                Letra_ColumnaExcel = "GV"
            Case 205
                Letra_ColumnaExcel = "GW"
            Case 206
                Letra_ColumnaExcel = "GX"
            Case 207
                Letra_ColumnaExcel = "GY"
            Case 208
                Letra_ColumnaExcel = "GZ"
            Case 209
                Letra_ColumnaExcel = "HA"
            Case 210
                Letra_ColumnaExcel = "HB"
            Case 211
                Letra_ColumnaExcel = "HC"
            Case 212
                Letra_ColumnaExcel = "HD"
            Case 213
                Letra_ColumnaExcel = "HE"
            Case 214
                Letra_ColumnaExcel = "HF"
            Case 215
                Letra_ColumnaExcel = "HG"
            Case 216
                Letra_ColumnaExcel = "HH"
            Case 217
                Letra_ColumnaExcel = "HI"
            Case 218
                Letra_ColumnaExcel = "HJ"
            Case 219
                Letra_ColumnaExcel = "HK"
            Case 220
                Letra_ColumnaExcel = "HL"
            Case 221
                Letra_ColumnaExcel = "HM"
            Case 222
                Letra_ColumnaExcel = "HN"
            Case 223
                Letra_ColumnaExcel = "HO"
            Case 224
                Letra_ColumnaExcel = "HP"
            Case 225
                Letra_ColumnaExcel = "HQ"
            Case 226
                Letra_ColumnaExcel = "HR"
            Case 227
                Letra_ColumnaExcel = "HS"
            Case 228
                Letra_ColumnaExcel = "HT"
            Case 229
                Letra_ColumnaExcel = "HU"
            Case 230
                Letra_ColumnaExcel = "HV"
            Case 231
                Letra_ColumnaExcel = "HW"
            Case 232
                Letra_ColumnaExcel = "HX"
            Case 233
                Letra_ColumnaExcel = "HY"
            Case 234
                Letra_ColumnaExcel = "HZ"
            Case 235
                Letra_ColumnaExcel = "IA"
            Case 236
                Letra_ColumnaExcel = "IB"
            Case 237
                Letra_ColumnaExcel = "IC"
            Case 238
                Letra_ColumnaExcel = "ID"
            Case 239
                Letra_ColumnaExcel = "IE"
            Case 240
                Letra_ColumnaExcel = "IF"
            Case 241
                Letra_ColumnaExcel = "IG"
            Case 242
                Letra_ColumnaExcel = "IH"
            Case 243
                Letra_ColumnaExcel = "II"
            Case 244
                Letra_ColumnaExcel = "IJ"
            Case 245
                Letra_ColumnaExcel = "IK"
            Case 246
                Letra_ColumnaExcel = "IL"
            Case 247
                Letra_ColumnaExcel = "IM"
            Case 248
                Letra_ColumnaExcel = "IN"
            Case 249
                Letra_ColumnaExcel = "IO"
            Case 250
                Letra_ColumnaExcel = "IP"
            Case 251
                Letra_ColumnaExcel = "IQ"
            Case 252
                Letra_ColumnaExcel = "IR"
            Case 253
                Letra_ColumnaExcel = "IS"
            Case 254
                Letra_ColumnaExcel = "IT"
            Case 255
                Letra_ColumnaExcel = "IU"
            Case 256
                Letra_ColumnaExcel = "IV"
            Case 257
                Letra_ColumnaExcel = "IW"
            Case 258
                Letra_ColumnaExcel = "IX"
            Case 259
                Letra_ColumnaExcel = "IY"
            Case 260
                Letra_ColumnaExcel = "IZ"
            Case 261
                Letra_ColumnaExcel = "JA"
            Case 262
                Letra_ColumnaExcel = "JB"
            Case 263
                Letra_ColumnaExcel = "JC"
            Case 264
                Letra_ColumnaExcel = "JD"
            Case 265
                Letra_ColumnaExcel = "JE"
            Case 266
                Letra_ColumnaExcel = "JF"
            Case 267
                Letra_ColumnaExcel = "JG"
            Case 268
                Letra_ColumnaExcel = "JH"
            Case 269
                Letra_ColumnaExcel = "JI"
            Case 270
                Letra_ColumnaExcel = "JJ"
            Case 271
                Letra_ColumnaExcel = "JK"
            Case 272
                Letra_ColumnaExcel = "JL"
            Case 273
                Letra_ColumnaExcel = "JM"
            Case 274
                Letra_ColumnaExcel = "JN"
            Case 275
                Letra_ColumnaExcel = "JO"
            Case 276
                Letra_ColumnaExcel = "JP"
            Case 277
                Letra_ColumnaExcel = "JQ"
            Case 278
                Letra_ColumnaExcel = "JR"
            Case 279
                Letra_ColumnaExcel = "JS"
            Case 280
                Letra_ColumnaExcel = "JT"
            Case 281
                Letra_ColumnaExcel = "JU"
            Case 282
                Letra_ColumnaExcel = "JV"
            Case 283
                Letra_ColumnaExcel = "JW"
            Case 284
                Letra_ColumnaExcel = "JX"
            Case 285
                Letra_ColumnaExcel = "JY"
            Case 286
                Letra_ColumnaExcel = "JZ"
            Case 287
                Letra_ColumnaExcel = "KA"
            Case 288
                Letra_ColumnaExcel = "KB"
            Case 289
                Letra_ColumnaExcel = "KC"
            Case 290
                Letra_ColumnaExcel = "KD"
            Case 291
                Letra_ColumnaExcel = "KE"
            Case 292
                Letra_ColumnaExcel = "KF"
            Case 293
                Letra_ColumnaExcel = "KG"
            Case 294
                Letra_ColumnaExcel = "KH"
            Case 295
                Letra_ColumnaExcel = "KI"
            Case 296
                Letra_ColumnaExcel = "KJ"
            Case 297
                Letra_ColumnaExcel = "KK"
            Case 298
                Letra_ColumnaExcel = "KL"
            Case 299
                Letra_ColumnaExcel = "KM"
            Case 300
                Letra_ColumnaExcel = "KN"
            Case 301
                Letra_ColumnaExcel = "KO"
            Case 302
                Letra_ColumnaExcel = "KP"
            Case 303
                Letra_ColumnaExcel = "KQ"
            Case 304
                Letra_ColumnaExcel = "KR"
            Case 305
                Letra_ColumnaExcel = "KS"
            Case 306
                Letra_ColumnaExcel = "KT"
            Case 307
                Letra_ColumnaExcel = "KU"
            Case 308
                Letra_ColumnaExcel = "KV"
            Case 309
                Letra_ColumnaExcel = "KW"
            Case 310
                Letra_ColumnaExcel = "KX"
            Case 311
                Letra_ColumnaExcel = "KY"
            Case 312
                Letra_ColumnaExcel = "KZ"
            Case 313
                Letra_ColumnaExcel = "LA"
            Case 314
                Letra_ColumnaExcel = "LB"
            Case 315
                Letra_ColumnaExcel = "LC"
            Case 316
                Letra_ColumnaExcel = "LD"
            Case 317
                Letra_ColumnaExcel = "LE"
            Case 318
                Letra_ColumnaExcel = "LF"
            Case 319
                Letra_ColumnaExcel = "LG"
            Case 320
                Letra_ColumnaExcel = "LH"
            Case 321
                Letra_ColumnaExcel = "LI"
            Case 322
                Letra_ColumnaExcel = "LJ"
            Case 323
                Letra_ColumnaExcel = "LK"
            Case 324
                Letra_ColumnaExcel = "LL"
            Case 325
                Letra_ColumnaExcel = "LM"
            Case 326
                Letra_ColumnaExcel = "LN"
            Case 327
                Letra_ColumnaExcel = "LO"
            Case 328
                Letra_ColumnaExcel = "LP"
            Case 329
                Letra_ColumnaExcel = "LQ"
            Case 330
                Letra_ColumnaExcel = "LR"
            Case 331
                Letra_ColumnaExcel = "LS"
            Case 332
                Letra_ColumnaExcel = "LT"
            Case 333
                Letra_ColumnaExcel = "LU"
            Case 334
                Letra_ColumnaExcel = "LV"
            Case 335
                Letra_ColumnaExcel = "LW"
            Case 336
                Letra_ColumnaExcel = "LX"
            Case 337
                Letra_ColumnaExcel = "LY"
            Case 338
                Letra_ColumnaExcel = "LZ"
            Case 339
                Letra_ColumnaExcel = "MA"
            Case 340
                Letra_ColumnaExcel = "MB"
            Case 341
                Letra_ColumnaExcel = "MC"
            Case 342
                Letra_ColumnaExcel = "MD"
            Case 343
                Letra_ColumnaExcel = "ME"
            Case 344
                Letra_ColumnaExcel = "MF"
            Case 345
                Letra_ColumnaExcel = "MG"
            Case 346
                Letra_ColumnaExcel = "MH"
            Case 347
                Letra_ColumnaExcel = "MI"
            Case 348
                Letra_ColumnaExcel = "MJ"
            Case 349
                Letra_ColumnaExcel = "MK"
            Case 350
                Letra_ColumnaExcel = "ML"
            Case 351
                Letra_ColumnaExcel = "MM"
            Case 352
                Letra_ColumnaExcel = "MN"
            Case 353
                Letra_ColumnaExcel = "MO"
            Case 354
                Letra_ColumnaExcel = "MP"
            Case 355
                Letra_ColumnaExcel = "MQ"
            Case 356
                Letra_ColumnaExcel = "MR"
            Case 357
                Letra_ColumnaExcel = "MS"
            Case 358
                Letra_ColumnaExcel = "MT"
            Case 359
                Letra_ColumnaExcel = "MU"
            Case 360
                Letra_ColumnaExcel = "MV"
            Case 361
                Letra_ColumnaExcel = "MW"
            Case 362
                Letra_ColumnaExcel = "MX"
            Case 363
                Letra_ColumnaExcel = "MY"
            Case 364
                Letra_ColumnaExcel = "MZ"
            Case 365
                Letra_ColumnaExcel = "NA"
            Case 366
                Letra_ColumnaExcel = "NB"
            Case 367
                Letra_ColumnaExcel = "NC"
            Case 368
                Letra_ColumnaExcel = "ND"
            Case 369
                Letra_ColumnaExcel = "NE"
            Case 370
                Letra_ColumnaExcel = "NF"
            Case 371
                Letra_ColumnaExcel = "NG"
            Case 372
                Letra_ColumnaExcel = "NH"
            Case 373
                Letra_ColumnaExcel = "NI"
            Case 374
                Letra_ColumnaExcel = "NJ"
            Case 375
                Letra_ColumnaExcel = "NK"
            Case 376
                Letra_ColumnaExcel = "NL"
            Case 377
                Letra_ColumnaExcel = "NM"
            Case 378
                Letra_ColumnaExcel = "NN"
            Case 379
                Letra_ColumnaExcel = "NO"
            Case 380
                Letra_ColumnaExcel = "NP"
            Case 381
                Letra_ColumnaExcel = "NQ"
            Case 382
                Letra_ColumnaExcel = "NR"
            Case 383
                Letra_ColumnaExcel = "NS"
            Case 384
                Letra_ColumnaExcel = "NT"
            Case 385
                Letra_ColumnaExcel = "NU"
            Case 386
                Letra_ColumnaExcel = "NV"
            Case 387
                Letra_ColumnaExcel = "NW"
            Case 388
                Letra_ColumnaExcel = "NX"
            Case 389
                Letra_ColumnaExcel = "NY"
            Case 390
                Letra_ColumnaExcel = "NZ"
            Case 391
                Letra_ColumnaExcel = "OA"
            Case 392
                Letra_ColumnaExcel = "OB"
            Case 393
                Letra_ColumnaExcel = "OC"
            Case 394
                Letra_ColumnaExcel = "OD"
            Case 395
                Letra_ColumnaExcel = "OE"
            Case 396
                Letra_ColumnaExcel = "OF"
            Case 397
                Letra_ColumnaExcel = "OG"
            Case 398
                Letra_ColumnaExcel = "OH"
            Case 399
                Letra_ColumnaExcel = "OI"
            Case 400
                Letra_ColumnaExcel = "OJ"
            Case 401
                Letra_ColumnaExcel = "OK"
            Case 402
                Letra_ColumnaExcel = "OL"
            Case 403
                Letra_ColumnaExcel = "OM"
            Case 404
                Letra_ColumnaExcel = "ON"
            Case 405
                Letra_ColumnaExcel = "OO"
            Case 406
                Letra_ColumnaExcel = "OP"
            Case 407
                Letra_ColumnaExcel = "OQ"
            Case 408
                Letra_ColumnaExcel = "OR"
            Case 409
                Letra_ColumnaExcel = "OS"
            Case 410
                Letra_ColumnaExcel = "OT"
            Case 411
                Letra_ColumnaExcel = "OU"
            Case 412
                Letra_ColumnaExcel = "OV"
            Case 413
                Letra_ColumnaExcel = "OW"
            Case 414
                Letra_ColumnaExcel = "OX"
            Case 415
                Letra_ColumnaExcel = "OY"
            Case 416
                Letra_ColumnaExcel = "OZ"
            Case 417
                Letra_ColumnaExcel = "PA"
            Case 418
                Letra_ColumnaExcel = "PB"
            Case 419
                Letra_ColumnaExcel = "PC"
            Case 420
                Letra_ColumnaExcel = "PD"
            Case 421
                Letra_ColumnaExcel = "PE"
            Case 422
                Letra_ColumnaExcel = "PF"
            Case 423
                Letra_ColumnaExcel = "PG"
            Case 424
                Letra_ColumnaExcel = "PH"
            Case 425
                Letra_ColumnaExcel = "PI"
            Case 426
                Letra_ColumnaExcel = "PJ"
            Case 427
                Letra_ColumnaExcel = "PK"
            Case 428
                Letra_ColumnaExcel = "PL"
            Case 429
                Letra_ColumnaExcel = "PM"
            Case 430
                Letra_ColumnaExcel = "PN"
            Case 431
                Letra_ColumnaExcel = "PO"
            Case 432
                Letra_ColumnaExcel = "PP"
            Case 433
                Letra_ColumnaExcel = "PQ"
            Case 434
                Letra_ColumnaExcel = "PR"
            Case 435
                Letra_ColumnaExcel = "PS"
            Case 436
                Letra_ColumnaExcel = "PT"
            Case 437
                Letra_ColumnaExcel = "PU"
            Case 438
                Letra_ColumnaExcel = "PV"
            Case 439
                Letra_ColumnaExcel = "PW"
            Case 440
                Letra_ColumnaExcel = "PX"
            Case 441
                Letra_ColumnaExcel = "PY"
            Case 442
                Letra_ColumnaExcel = "PZ"
            Case 443
                Letra_ColumnaExcel = "QA"
            Case 444
                Letra_ColumnaExcel = "QB"
            Case 445
                Letra_ColumnaExcel = "QC"
            Case 446
                Letra_ColumnaExcel = "QD"
            Case 447
                Letra_ColumnaExcel = "QE"
            Case 448
                Letra_ColumnaExcel = "QF"
            Case 449
                Letra_ColumnaExcel = "QG"
            Case 450
                Letra_ColumnaExcel = "QH"
            Case 451
                Letra_ColumnaExcel = "QI"
            Case 452
                Letra_ColumnaExcel = "QJ"
            Case 453
                Letra_ColumnaExcel = "QK"
            Case 454
                Letra_ColumnaExcel = "QL"
            Case 455
                Letra_ColumnaExcel = "QM"
            Case 456
                Letra_ColumnaExcel = "QN"
            Case 457
                Letra_ColumnaExcel = "QO"
            Case 458
                Letra_ColumnaExcel = "QP"
            Case 459
                Letra_ColumnaExcel = "QQ"
            Case 460
                Letra_ColumnaExcel = "QR"
            Case 461
                Letra_ColumnaExcel = "QS"
            Case 462
                Letra_ColumnaExcel = "QT"
            Case 463
                Letra_ColumnaExcel = "QU"
            Case 464
                Letra_ColumnaExcel = "QV"
            Case 465
                Letra_ColumnaExcel = "QW"
            Case 466
                Letra_ColumnaExcel = "QX"
            Case 467
                Letra_ColumnaExcel = "QY"
            Case 468
                Letra_ColumnaExcel = "QZ"
            Case 469
                Letra_ColumnaExcel = "RA"
            Case 470
                Letra_ColumnaExcel = "RB"
            Case 471
                Letra_ColumnaExcel = "RC"
            Case 472
                Letra_ColumnaExcel = "RD"
            Case 473
                Letra_ColumnaExcel = "RE"
            Case 474
                Letra_ColumnaExcel = "RF"
            Case 475
                Letra_ColumnaExcel = "RG"
            Case 476
                Letra_ColumnaExcel = "RH"
            Case 477
                Letra_ColumnaExcel = "RI"
            Case 478
                Letra_ColumnaExcel = "RJ"
            Case 479
                Letra_ColumnaExcel = "RK"
            Case 480
                Letra_ColumnaExcel = "RL"
            Case 481
                Letra_ColumnaExcel = "RM"
            Case 482
                Letra_ColumnaExcel = "RN"
            Case 483
                Letra_ColumnaExcel = "RO"
            Case 484
                Letra_ColumnaExcel = "RP"
            Case 485
                Letra_ColumnaExcel = "RQ"
            Case 486
                Letra_ColumnaExcel = "RR"
            Case 487
                Letra_ColumnaExcel = "RS"
            Case 488
                Letra_ColumnaExcel = "RT"
            Case 489
                Letra_ColumnaExcel = "RU"
            Case 490
                Letra_ColumnaExcel = "RV"
            Case 491
                Letra_ColumnaExcel = "RW"
            Case 492
                Letra_ColumnaExcel = "RX"
            Case 493
                Letra_ColumnaExcel = "RY"
            Case 494
                Letra_ColumnaExcel = "RZ"
            Case 495
                Letra_ColumnaExcel = "SA"
            Case 496
                Letra_ColumnaExcel = "SB"
            Case 497
                Letra_ColumnaExcel = "SC"
            Case 498
                Letra_ColumnaExcel = "SD"
            Case 499
                Letra_ColumnaExcel = "SE"
            Case 500
                Letra_ColumnaExcel = "SF"
            Case 501
                Letra_ColumnaExcel = "SG"
            Case 502
                Letra_ColumnaExcel = "SH"
            Case 503
                Letra_ColumnaExcel = "SI"
            Case 504
                Letra_ColumnaExcel = "SJ"
            Case 505
                Letra_ColumnaExcel = "SK"
            Case 506
                Letra_ColumnaExcel = "SL"
            Case 507
                Letra_ColumnaExcel = "SM"
            Case 508
                Letra_ColumnaExcel = "SN"
            Case 509
                Letra_ColumnaExcel = "SO"
            Case 510
                Letra_ColumnaExcel = "SP"
            Case 511
                Letra_ColumnaExcel = "SQ"
            Case 512
                Letra_ColumnaExcel = "SR"
            Case 513
                Letra_ColumnaExcel = "SS"
            Case 514
                Letra_ColumnaExcel = "ST"
            Case 515
                Letra_ColumnaExcel = "SU"
            Case 516
                Letra_ColumnaExcel = "SV"
            Case 517
                Letra_ColumnaExcel = "SW"
            Case 518
                Letra_ColumnaExcel = "SX"
            Case 519
                Letra_ColumnaExcel = "SY"
            Case 520
                Letra_ColumnaExcel = "SZ"
            Case 521
                Letra_ColumnaExcel = "TA"
            Case 522
                Letra_ColumnaExcel = "TB"
            Case 523
                Letra_ColumnaExcel = "TC"
            Case 524
                Letra_ColumnaExcel = "TD"
            Case 525
                Letra_ColumnaExcel = "TE"
            Case 526
                Letra_ColumnaExcel = "TF"
            Case 527
                Letra_ColumnaExcel = "TG"
            Case 528
                Letra_ColumnaExcel = "TH"
            Case 529
                Letra_ColumnaExcel = "TI"
            Case 530
                Letra_ColumnaExcel = "TJ"
            Case 531
                Letra_ColumnaExcel = "TK"
            Case 532
                Letra_ColumnaExcel = "TL"
            Case 533
                Letra_ColumnaExcel = "TM"
            Case 534
                Letra_ColumnaExcel = "TN"
            Case 535
                Letra_ColumnaExcel = "TO"
            Case 536
                Letra_ColumnaExcel = "TP"
            Case 537
                Letra_ColumnaExcel = "TQ"
            Case 538
                Letra_ColumnaExcel = "TR"
            Case 539
                Letra_ColumnaExcel = "TS"
            Case 540
                Letra_ColumnaExcel = "TT"
            Case 541
                Letra_ColumnaExcel = "TU"
            Case 542
                Letra_ColumnaExcel = "TV"
            Case 543
                Letra_ColumnaExcel = "TW"
            Case 544
                Letra_ColumnaExcel = "TX"
            Case 545
                Letra_ColumnaExcel = "TY"
            Case 546
                Letra_ColumnaExcel = "TZ"
            Case Else
                Letra_ColumnaExcel = ""
        End Select
    End Function

Public Function fc_Numero(pMes As Integer) As String
Dim xMes As String
Select Case pMes
Case 1: xMes = "UN "
Case 2: xMes = "DOS "
Case 3: xMes = "TRES "
Case 4: xMes = "CUATRO "
Case 5: xMes = "CINCO"
Case 6: xMes = "SEIS "
Case 7: xMes = "SIETE "
Case 8: xMes = "OCHO "
Case 9: xMes = "NUEVE "
Case 10: xMes = "DIEX "
Case 11: xMes = "ONCE "
Case 12: xMes = "DOCE "
Case 13: xMes = "TRECE "
Case 14: xMes = "CATORCE "
Case 15: xMes = "QUINCE "
Case 16: xMes = "DIECISEIS "
Case 17: xMes = "DIECISIETE "
Case 18: xMes = "DIECIOCHO "
Case 19: xMes = "DIECINUEVE "
Case 20: xMes = "VEINTE "
Case 21: xMes = "VEINTIUNO "
Case 22: xMes = "VEINTIDOS "
Case 23: xMes = "VEINTITRES "
Case 24: xMes = "VEINTICUATRO "
Case 36: xMes = "TREINA Y SEIS "
Case 48: xMes = "CUARENTA Y OCHO "
Case Else
    xMes = ""
End Select
fc_Numero = xMes
End Function

Sub GenerarAsientoProvision(ByVal Periodo As String, ByVal tipoprovision As String, ByVal TipoTrabajador As String, ByVal lote As String, ByVal Voucher As String, ByVal titulo As String, ByVal Mes As String, ByVal SubDiario As String)
Dim m As Long
On Error GoTo GestionaErrores

Dim cmd As New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "spAsientoProvision"
cmd.ActiveConnection = cn
cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, , 0)
cmd.Parameters.Append cmd.CreateParameter("@cia", adChar, adParamInput, 2, wcia)
cmd.Parameters.Append cmd.CreateParameter("@periodo", adChar, adParamInput, 6, Periodo)
cmd.Parameters.Append cmd.CreateParameter("@tipoprovision", adChar, adParamInput, 2, tipoprovision)
cmd.Parameters.Append cmd.CreateParameter("@tipotrabajador", adChar, adParamInput, 2, TipoTrabajador)
cmd.Parameters.Append cmd.CreateParameter("@usuario", adVarChar, adParamInput, 100, wuser)
cmd.Parameters.Append cmd.CreateParameter("@pc", adVarChar, adParamInput, 100, wNamePC)
cmd.CommandTimeout = 0
Screen.MousePointer = 0
cmd.Execute
Dim resultado As Integer
resultado = cmd.Parameters("Return").Value
Select Case resultado
Case 0
    MsgBox "No se tienen registrado el calculo de provisión, efectue primero el calculo ", vbCritical
Case 1
    MsgBox "Asiento emitido satisfactoriamente ", vbInformation
    
    Call Carga_Asiento_Excel(Left(Periodo, 4), Val(Right(Periodo, 2)), lote, Voucher, titulo, Mes, "S")
Case 2
    MsgBox "Asiento inconsistente, revisar seteo de centros de costo ", vbCritical
Case Else
    MsgBox "No es posible generar asiento de provisión, vuelva a intentarlo luego ", vbCritical
End Select

Set cmd = Nothing

GestionaErrores:
m = Err.Number

Select Case m

        Case Is = -2147217843
             MsgBox "Usuario  y/o Clave de Acceso Incorrectos", vbCritical
             Exit Sub
        Case Is = -2147467259
             MsgBox "No hay Coneccion con el Servidor", vbCritical
             Exit Sub
        Case Is = 0
        Case Else
              MsgBox Err.Description, vbCritical, "Error " & Err.Number
              Exit Sub
 End Select


End Sub

Sub Procesa_Archivo_Banco_Excel(mano As Integer, mmes As Integer)
Dim March As String
Dim MarchOri As String
Dim nFil As Integer
Dim lDesTipoBol As String

Dim lCta As String

Dim lMon As String
If moneda = "S/." Then lMon = "SOL" Else lMon = "DOL"

'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object

Sql$ = "uSP_TRAE_DEPO_CTS '" & wcia & "'," & mano & "," & mmes & ""

If Not (fAbrRst(rs2, Sql$)) Then rs2.Close: Set rs2 = Nothing: Exit Sub
rs2.MoveFirst

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application

xlApp2.Sheets.Add
xlApp2.Sheets("Hoja1").Select
xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
xlApp2.Sheets("Hoja2").Select
xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)


Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("Hoja1")
xlSheet.Name = "Soles"

xlSheet.Range("B:B").NumberFormat = "@"
xlSheet.Range("H:H").NumberFormat = "@"

Dim I As Integer

xlSheet.Activate

xlSheet.Range("A:A").ColumnWidth = 5
xlSheet.Range("B:B").ColumnWidth = 12
xlSheet.Range("C:C").ColumnWidth = 8
xlSheet.Range("D:F").ColumnWidth = 20
xlSheet.Range("Q:Q").ColumnWidth = 11
xlSheet.Range("H:H").ColumnWidth = 20
xlSheet.Range("I:I").ColumnWidth = 12
xlSheet.Range("J:J").ColumnWidth = 9
xlSheet.Range("K:K").ColumnWidth = 16
xlSheet.Range("G:G").HorizontalAlignment = xlCenter

xlSheet.Cells(1, 1).Value = Trae_CIA(wcia)
xlSheet.Cells(3, 1).Value = "DEPOSITO CTS " & Name_Month(Format(mmes, "00")) & " - " & Format(mano, "0000")
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).Merge

xlSheet.Cells(5, 1).Value = "ITEM"
xlSheet.Cells(5, 2).Value = "DNI"
xlSheet.Cells(5, 3).Value = "IND.DNI"
xlSheet.Cells(5, 4).Value = "APE.PATERNO"
xlSheet.Cells(5, 5).Value = "AP.MATERNO"
xlSheet.Cells(5, 6).Value = "NOMBRES"
xlSheet.Cells(5, 7).Value = "FEC.NAC."
xlSheet.Cells(5, 8).Value = "NRO CTA"
xlSheet.Cells(5, 9).Value = "DEPOSITAR"
xlSheet.Cells(5, 10).Value = "MONEDA"
xlSheet.Cells(5, 11).Value = "REMUNERACIONES"

xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).Font.Bold = True

nFil = 5
Dim lItem As Integer
lItem = 1
If rs2.RecordCount > 0 Then rs2.MoveFirst
Dim lBanco As String
lBanco = ""
Do While Not rs2.EOF
   If Trim(rs2!ctsmoneda & "") = "S/." Then
      If lBanco <> Trim(rs2!desbanco & "") Then
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs2!desbanco & "")
         xlSheet.Cells(nFil, 2).Font.Bold = True
         nFil = nFil + 2
         lBanco = Trim(rs2!desbanco & "")
         lItem = 1
      End If
      xlSheet.Cells(nFil, 1).Value = lItem
      If Trim(rs2!td & "") = "01" Then
         xlSheet.Cells(nFil, 2).Value = Trim(rs2!nd & "")
         xlSheet.Cells(nFil, 3).Value = 1
      End If
      xlSheet.Cells(nFil, 4).Value = Trim(rs2!apep & "")
      xlSheet.Cells(nFil, 5).Value = Trim(rs2!apem & "")
      xlSheet.Cells(nFil, 6).Value = Trim(rs2!NOM & "")
      xlSheet.Cells(nFil, 7).Value = Format(rs2!dianac, "00") & "/" & Format(rs2!mesnac, "00") & "/" & Format(rs2!anonac, "0000")
      xlSheet.Cells(nFil, 8).Value = Trim(rs2!Cuenta & "")
      xlSheet.Cells(nFil, 9).Value = rs2!totneto
      xlSheet.Cells(nFil, 10).Value = "Soles"
      xlSheet.Cells(nFil, 11).Value = rs2!Remun
      nFil = nFil + 1
      lItem = lItem + 1
   End If
   rs2.MoveNext
Loop

'DOLARES
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "Dolares"

xlSheet.Range("B:B").NumberFormat = "@"
xlSheet.Range("H:H").NumberFormat = "@"

xlSheet.Activate

xlSheet.Range("A:A").ColumnWidth = 5
xlSheet.Range("B:B").ColumnWidth = 12
xlSheet.Range("C:C").ColumnWidth = 8
xlSheet.Range("D:F").ColumnWidth = 20
xlSheet.Range("Q:Q").ColumnWidth = 11
xlSheet.Range("H:H").ColumnWidth = 20
xlSheet.Range("I:I").ColumnWidth = 12
xlSheet.Range("J:J").ColumnWidth = 9
xlSheet.Range("K:K").ColumnWidth = 16
xlSheet.Range("G:G").HorizontalAlignment = xlCenter

xlSheet.Cells(1, 1).Value = Trae_CIA(wcia)
xlSheet.Cells(3, 1).Value = "DEPOSITO CTS " & Name_Month(Format(mmes, "00")) & " - " & Format(mano, "0000")
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).Merge

xlSheet.Cells(5, 1).Value = "ITEM"
xlSheet.Cells(5, 2).Value = "DNI"
xlSheet.Cells(5, 3).Value = "IND.DNI"
xlSheet.Cells(5, 4).Value = "APE.PATERNO"
xlSheet.Cells(5, 5).Value = "AP.MATERNO"
xlSheet.Cells(5, 6).Value = "NOMBRES"
xlSheet.Cells(5, 7).Value = "FEC.NAC."
xlSheet.Cells(5, 8).Value = "NRO CTA"
xlSheet.Cells(5, 9).Value = "DEPOSITAR"
xlSheet.Cells(5, 10).Value = "MONEDA"
xlSheet.Cells(5, 11).Value = "REMUNERACIONES"

xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).Font.Bold = True

nFil = 5
lItem = 1
If rs2.RecordCount > 0 Then rs2.MoveFirst
lBanco = ""
Do While Not rs2.EOF
   If Trim(rs2!ctsmoneda & "") = "US$" Then
      If lBanco <> Trim(rs2!desbanco & "") Then
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs2!desbanco & "")
         xlSheet.Cells(nFil, 2).Font.Bold = True
         nFil = nFil + 2
         lBanco = Trim(rs2!desbanco & "")
         lItem = 1
      End If
      xlSheet.Cells(nFil, 1).Value = lItem
      If Trim(rs2!td & "") = "01" Then
         xlSheet.Cells(nFil, 2).Value = Trim(rs2!nd & "")
         xlSheet.Cells(nFil, 3).Value = 1
      End If
      xlSheet.Cells(nFil, 4).Value = Trim(rs2!apep & "")
      xlSheet.Cells(nFil, 5).Value = Trim(rs2!apem & "")
      xlSheet.Cells(nFil, 6).Value = Trim(rs2!NOM & "")
      xlSheet.Cells(nFil, 7).Value = Format(rs2!dianac, "00") & "/" & Format(rs2!mesnac, "00") & "/" & Format(rs2!anonac, "0000")
      xlSheet.Cells(nFil, 8).Value = Trim(rs2!Cuenta & "")
      xlSheet.Cells(nFil, 9).Value = rs2!Deposito
      xlSheet.Cells(nFil, 10).Value = "Soles"
      xlSheet.Cells(nFil, 11).Value = rs2!Remun
      xlSheet.Cells(nFil, 14).Value = rs2!totneto
      nFil = nFil + 1
      lItem = lItem + 1
   End If
   rs2.MoveNext
Loop


For I = 1 To 3
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridlines = False
Next
xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Dim lfPRoc As String
lfPRoc = ""
Sql = "SELECT * FROM Pla_Tc_Cts Where cia='" & wcia & "' and ano=" & mano & " and mes=" & mmes & " and status<>'*'"
If (fAbrRst(Rs, Sql)) Then
   lfPRoc = Str(Year(Rs!FecDepo)) + Format(Month(Rs!FecDepo), "00") + Format(Day(Rs!FecDepo), "00")
End If
Rs.Close: Set Rs = Nothing

Crea_Rs_Depo_Cts
Call Print_TxtBcoCredito_CTS("S/.", lfPRoc)
Call Print_TxtBcoCredito_CTS("US$", lfPRoc)

Call Print_TxtBcoConti_CTS("S/.", lfPRoc)
Call Print_TxtBcoConti_4UltBol("S/.", lfPRoc)
Call Print_TxtBcoConti_CTS("US$", lfPRoc)
Call Print_TxtBcoConti_4UltBol("US$", lfPRoc)

Call Print_TxtBcoScot_CTS("S/.", lfPRoc)
Call Print_TxtBcoScot_CTS("US$", lfPRoc)

Call Print_TxtBcoCUZCO_CTS("S/.", lfPRoc)
Call Print_TxtBcoCUZCO_CTS("US$", lfPRoc)

End Sub
Private Sub Print_TxtBcoCredito_CTS(lMon As String, mfproceso As String)

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double


If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim xMon As String
If lMon = "S/." Then xMon = "0001" Else xMon = "1001"

If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
    If Trim(rs2!ctsbanco & "") = "01" And Trim(rs2!ctsmoneda & "") = lMon Then
        rsdepo.AddNew
        rsdepo!Codigo = Trim(rs2!PlaCod & "")
        rsdepo!NOM_CLIE = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
        rsdepo!NUMEROCTA = Mid(rs2!Cuenta, 1, 3) + Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        
        rsdepo!importe = rs2!Deposito
        rsdepo!TIPO_REG = "2"
        rsdepo!Cuenta = "A"
        If Trim(rs2!td & "") = "01" Then
            rsdepo!tipo_doc = "1"
            rsdepo!NUMERO_DOC = Trim(rs2!nd & "")
        End If
        
        rsdepo!REF_TRAB = "DEPOSITO CTS"
        rsdepo!REF_EMP = "DEPOSITO CTS"
        rsdepo!moneda = xMon
        rsdepo!FLAG = "S"
        
        rsdepo!cta = Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        'rsdepo!CTAINTER = Trim(rsbco!CTAINTER & "")
        'rsdepo!Sucursal = Trim(rsbco!Sucursal & "")
        'rsdepo!PAGOCUENTA = Trim(rsbco!PAGOCUENTA & "")
        rsdepo!Remunera = rs2!Remun
        rsdepo!ctscci = Trim(rs2!ctscci & "")
    End If
    rs2.MoveNext
Loop


nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!NUMEROCTA & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   nTrab = nTrab + 1
   mscta = mscta + Val(rsdepo!cta)
   If Trim(rsdepo!importe & "") <> "" Then mneto = mneto + rsdepo!importe
  rsdepo.MoveNext
Loop

If lMon = "S/." Then
   RUTA$ = App.Path & "\REPORTS\CTSCredSol.txt"
Else
   RUTA$ = App.Path & "\REPORTS\CTSCredDol.txt"
End If
Open RUTA$ For Output As #1

Dim mcad As String
Dim MCOD_CTA As Double

mcad = "1" & Llenar_Ceros(Trim(Str(nTrab)), 6) & Trim(mfproceso)
If lMon = "S/." Then
   mcad = mcad & "C00011910215470064       620100037689 " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
   MCOD_CTA = 215470064
Else
   mcad = mcad & "C10011910686330135       620100037689 " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
   MCOD_CTA = 686330
End If

mcad = mcad & "Deposito CTS        " & Space(20)

mscta = mscta + MCOD_CTA

mcad = mcad & Llenar_Ceros(Trim(Str(mscta)), 15)
Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   'If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
   If rsdepo!importe > 0 Then
      Print #1, rsdepo!TIPO_REG + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + "0001" + Llenar_Ceros(Trim(fCadNum(rsdepo!Remunera, "#######0.00")), 17)
   End If
   rsdepo.MoveNext
 Loop
Close #1
End Sub
Private Sub Print_TxtBcoCUZCO_CTS(lMon As String, mfproceso As String)

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double


If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim xMon As String
If lMon = "S/." Then xMon = "0001" Else xMon = "1001"

If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
    If Trim(rs2!ctsbanco & "") = "71" And Trim(rs2!ctsmoneda & "") = lMon Then
        rsdepo.AddNew
        rsdepo!Codigo = Trim(rs2!PlaCod & "")
        rsdepo!NOM_CLIE = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
        rsdepo!NUMEROCTA = Mid(rs2!Cuenta, 1, 3) + Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        
        rsdepo!importe = rs2!Deposito
        rsdepo!TIPO_REG = "2"
        rsdepo!Cuenta = "A"
        If Trim(rs2!td & "") = "01" Then
            rsdepo!tipo_doc = "1"
            rsdepo!NUMERO_DOC = Trim(rs2!nd & "")
        End If
        
        rsdepo!REF_TRAB = "DEPOSITO CTS"
        rsdepo!REF_EMP = "DEPOSITO CTS"
        rsdepo!moneda = xMon
        rsdepo!FLAG = "S"
        
        rsdepo!cta = Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        'rsdepo!CTAINTER = Trim(rsbco!CTAINTER & "")
        'rsdepo!Sucursal = Trim(rsbco!Sucursal & "")
        'rsdepo!PAGOCUENTA = Trim(rsbco!PAGOCUENTA & "")
        rsdepo!Remunera = rs2!Remun
        rsdepo!ctscci = Trim(rs2!ctscci & "")
    End If
    rs2.MoveNext
Loop


nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!NUMEROCTA & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   nTrab = nTrab + 1
   mscta = mscta + Val(rsdepo!cta)
   If Trim(rsdepo!importe & "") <> "" Then mneto = mneto + rsdepo!importe
  rsdepo.MoveNext
Loop

If lMon = "S/." Then
   RUTA$ = App.Path & "\REPORTS\CTSCUZCOSol.txt"
Else
   RUTA$ = App.Path & "\REPORTS\CTSCUZCODol.txt"
End If
Open RUTA$ For Output As #1

Dim mcad As String
Dim MCOD_CTA As Double

mcad = "1" & Llenar_Ceros(Trim(Str(nTrab)), 6) & Trim(mfproceso)
If lMon = "S/." Then
   mcad = mcad & "C00011910215470064       620100037689 " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
   MCOD_CTA = 215470064
Else
   mcad = mcad & "C10011910686330135       620100037689 " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
   MCOD_CTA = 686330
End If

mcad = mcad & "Deposito CTS        " & Space(20)

mscta = mscta + MCOD_CTA

mcad = mcad & Llenar_Ceros(Trim(Str(mscta)), 15)
Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   'If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
   If rsdepo!importe > 0 Then
      Print #1, rsdepo!TIPO_REG + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + "0001" + Llenar_Ceros(Trim(fCadNum(rsdepo!Remunera, "#######0.00")), 17)
   End If
   rsdepo.MoveNext
 Loop
Close #1
End Sub

Private Sub Print_TxtBcoConti_CTS(lMon As String, mfproceso As String)

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double
Dim xCtaCom As String

mfproceso = Str(Year(Date)) + Format(Month(Date), "00") + Format(Day(Date), "00")

If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim xMon As String
Dim llm As String
If lMon = "S/." Then
   xMon = "600"
   llm = "PEN"
   xCtaCom = "00110686000100006678"
Else
   xMon = "610"
   llm = "USD"
   xCtaCom = "00110686000100006805"
End If
    
If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
    If Trim(rs2!ctsbanco & "") = "02" And Trim(rs2!ctsmoneda & "") = lMon Then
        rsdepo.AddNew
        rsdepo!Codigo = Trim(rs2!PlaCod & "")
        rsdepo!NOM_CLIE = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
        rsdepo!NUMEROCTA = Trim(rs2!Cuenta & "")
        
        rsdepo!importe = rs2!Deposito * 100
        rsdepo!TIPO_REG = "2"
        rsdepo!Cuenta = "A"
        If Trim(rs2!td & "") = "01" Then
            rsdepo!tipo_doc = "L"
            rsdepo!NUMERO_DOC = Trim(rs2!nd & "")
        End If
        
        rsdepo!REF_TRAB = "DEPOSITO CTS"
        rsdepo!REF_EMP = "DEPOSITO CTS"
        rsdepo!moneda = xMon
        rsdepo!FLAG = "S"
        
        rsdepo!cta = Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        'rsdepo!CTAINTER = Trim(rsbco!CTAINTER & "")
        'rsdepo!Sucursal = Trim(rsbco!Sucursal & "")
        'rsdepo!PAGOCUENTA = Trim(rsbco!PAGOCUENTA & "")
        rsdepo!Remunera = rs2!Remun * 100
        rsdepo!ctscci = Trim(rs2!ctscci & "")
    End If
    rs2.MoveNext
Loop


nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!NUMEROCTA & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   nTrab = nTrab + 1
   mscta = mscta + Val(rsdepo!cta)
   If Trim(rsdepo!importe & "") <> "" Then mneto = mneto + (rsdepo!importe / 100)
  rsdepo.MoveNext
Loop

If lMon = "S/." Then
   RUTA$ = App.Path & "\REPORTS\CTSContiSol.txt"
Else
   RUTA$ = App.Path & "\REPORTS\CTSContiDol.txt"
End If
Open RUTA$ For Output As #1

Dim mcad As String
Dim MCOD_CTA As Double

mneto = mneto * 100
mcad = xMon & xCtaCom & llm & Llenar_Ceros(Trim(fCadNum(mneto, "############000")), 15) & "F" & Trim(mfproceso) & "D" & "DEPOSITO CTS             " & Llenar_Ceros(Trim(Str(nTrab)), 6) & "S" & Space(68)
Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   'If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
   If rsdepo!importe > 0 Then
      'Print #1, rsdepo!TIPO_REG + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + "0001" + Llenar_Ceros(Trim(fCadNum(rsdepo!remunera, "#######0.00")), 17)
      Print #1, "002" + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + "P" + rsdepo!NUMEROCTA + Mid(rsdepo!NOM_CLIE, 1, 40) + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "############000")), 15) + "DEPOSITO CTS" + Space(40 - 12) + Space(1) + Space(50); Space(53) + Llenar_Ceros(Trim(fCadNum(rsdepo!Remunera, "############000")), 15) & "PEN" + Space(18)
   End If
   rsdepo.MoveNext
 Loop
Close #1
End Sub
Private Sub Print_TxtBcoConti_4UltBol(lMon As String, mfproceso As String)

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double
Dim xCtaCom As String

mfproceso = Str(Year(Date)) + Format(Month(Date), "00") + Format(Day(Date), "00")

If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim xMon As String
Dim llm As String
If lMon = "S/." Then
   xMon = "620"
   llm = "PEN"
   xCtaCom = "                    "
Else
   xMon = "620"
   llm = "USD"
   xCtaCom = "                    "
End If
    
If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
    If Trim(rs2!ctsbanco & "") = "02" And Trim(rs2!ctsmoneda & "") = lMon Then
        rsdepo.AddNew
        rsdepo!Codigo = Trim(rs2!PlaCod & "")
        rsdepo!NOM_CLIE = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
        rsdepo!NUMEROCTA = Trim(rs2!Cuenta & "")

        rsdepo!importe = rs2!Deposito * 100
        rsdepo!TIPO_REG = "2"
        rsdepo!Cuenta = "A"
        If Trim(rs2!td & "") = "01" Then
            rsdepo!tipo_doc = "L"
            rsdepo!NUMERO_DOC = Trim(rs2!nd & "")
        End If
        
        rsdepo!REF_TRAB = "remunePEN"
        rsdepo!REF_EMP = "remunePEN"
        rsdepo!moneda = xMon
        rsdepo!FLAG = "S"
        
        rsdepo!cta = Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        rsdepo!Remunera = rs2!Remun * 100
        rsdepo!ctscci = Trim(rs2!ctscci & "")
     End If
    rs2.MoveNext
Loop


nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!NUMEROCTA & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   nTrab = nTrab + 1
   mscta = mscta + Val(rsdepo!cta)
   If Trim(rsdepo!importe & "") <> "" Then mneto = mneto + (rsdepo!importe / 100)
  rsdepo.MoveNext
Loop

If lMon = "S/." Then
   RUTA$ = App.Path & "\REPORTS\BBVA_REM_CTS_SOL.txt"
Else
   RUTA$ = App.Path & "\REPORTS\BBVA_REM_CTS_DOL.txt"
End If
Open RUTA$ For Output As #1

Dim mcad As String
Dim MCOD_CTA As Double

mneto = mneto * 100
If lMon = "S/." Then
mcad = xMon & xCtaCom & llm & "               " & "F" & Trim(mfproceso) & "C4" & "remunePEN                " & Llenar_Ceros(Trim(Str(nTrab)), 6) & "S" & Space(68)
Else
mcad = xMon & xCtaCom & llm & "               " & "F" & Trim(mfproceso) & "C4" & "remuneUSD                " & Llenar_Ceros(Trim(Str(nTrab)), 6) & "S" & Space(68)
End If

Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   'If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
   If rsdepo!importe > 0 Then
      'Print #1, rsdepo!TIPO_REG + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + "0001" + Llenar_Ceros(Trim(fCadNum(rsdepo!remunera, "#######0.00")), 17)
      Print #1, "002" + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + "P" + rsdepo!NUMEROCTA + Mid(rsdepo!NOM_CLIE, 1, 40) + Llenar_Ceros(Trim(fCadNum(rsdepo!Remunera, "############000")), 15) + "PEN         "
      'Print #1, "002" + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + "P" + rsdepo!NUMEROCTA + Mid(rsdepo!NOM_CLIE, 1, 40) + Llenar_Ceros(Trim(fCadNum(rsdepo!Remunera, "############000")), 15) + "PEN         " + Space(40 - 12) + Space(1) + Space(50); Space(53) + Llenar_Ceros(Trim(fCadNum(rsdepo!Remunera, "############000")), 15) & "PEN" + Space(18)
   End If
   rsdepo.MoveNext
 Loop
Close #1
End Sub
Private Sub Print_TxtBcoScot_CTS(lMon As String, mfproceso As String)

If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim xMon As String
Dim llm As String
If lMon = "S/." Then
   xMon = "620"
   llm = "PEN"
Else
   xMon = "630"
   llm = "USD"
End If
    
If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
    If Trim(rs2!ctsbanco & "") = "29" And Trim(rs2!ctsmoneda & "") = lMon Then
        rsdepo.AddNew
        rsdepo!Codigo = Trim(rs2!PlaCod & "")
        rsdepo!NOM_CLIE = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
        rsdepo!NUMEROCTA = Mid(Trim(rs2!Cuenta & ""), 1, 3) + Mid(Trim(rs2!Cuenta & ""), 5, 7)
        
        rsdepo!importe = rs2!Deposito * 100
        rsdepo!TIPO_REG = "2"
        rsdepo!Cuenta = "A"
        If Trim(rs2!td & "") = "01" Then
            rsdepo!tipo_doc = "L"
            rsdepo!NUMERO_DOC = Trim(rs2!nd & "")
        Else
            rsdepo!NUMERO_DOC = Trim(rs2!PlaCod & "")
        End If
        
        rsdepo!REF_TRAB = "DEPOSITO CTS"
        rsdepo!REF_EMP = "DEPOSITO CTS"
        rsdepo!moneda = xMon
        rsdepo!FLAG = "S"
        
        rsdepo!cta = Mid(rs2!Cuenta, 5, 8) + Mid(rs2!Cuenta, 14, 1) + Mid(rs2!Cuenta, 16, 2)
        'rsdepo!CTAINTER = Trim(rsbco!CTAINTER & "")
        'rsdepo!Sucursal = Trim(rsbco!Sucursal & "")
        'rsdepo!PAGOCUENTA = Trim(rsbco!PAGOCUENTA & "")
        rsdepo!Remunera = rs2!Remun
        rsdepo!ctscci = Trim(rs2!ctscci & "")
    End If
    rs2.MoveNext
Loop


If lMon = "S/." Then
   RUTA$ = App.Path & "\REPORTS\CTSScotSol.txt"
Else
   RUTA$ = App.Path & "\REPORTS\CTSScotDol.txt"
End If
Open RUTA$ For Output As #1

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   'If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
   If rsdepo!importe > 0 Then
      Print #1, Mid(rsdepo!NUMERO_DOC, 1, 8) + Mid(rsdepo!NOM_CLIE, 1, 30) + "DEPOSITO DE CTS    " + mfproceso + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######000")), 11) + "3" + Mid(rsdepo!NUMEROCTA, 1, 10) + Mid(rsdepo!NUMERO_DOC, 1, 8) + " " + rsdepo!ctscci
   End If
   rsdepo.MoveNext
 Loop
Close #1
End Sub

Private Sub Crea_Rs_Depo_Cts()
    
    If rsdepo.State = 1 Then rsdepo.Close
    rsdepo.Fields.Append "TIPO_REG", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "CUENTA", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "NUMEROCTA", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "TIPO_DOC", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "NUMERO_DOC", adChar, 12, adFldIsNullable
    rsdepo.Fields.Append "NOM_CLIE", adChar, 75, adFldIsNullable
    rsdepo.Fields.Append "REF_TRAB", adChar, 40, adFldIsNullable
    rsdepo.Fields.Append "REF_EMP", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "MONEDA", adChar, 4, adFldIsNullable
    rsdepo.Fields.Append "IMPORTE", adDouble, 2, adFldIsNullable
    rsdepo.Fields.Append "FLAG", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "cta", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "CODIGO", adChar, 8, adFldIsNullable
    rsdepo.Fields.Append "CTAINTER", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "SUCURSAL", adChar, 3, adFldIsNullable
    rsdepo.Fields.Append "PAGOCUENTA", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "REMUNERA", adDouble, 2, adFldIsNullable
    rsdepo.Fields.Append "CTSCCI", adChar, 20, adFldIsNullable
   ' rsdepo.Fields.Append "TOTNETO", adDouble, 2, adFldIsNullable
    rsdepo.Open

End Sub


Public Function Valida_Cuenta(cta As String) As Boolean
Dim Rq As ADODB.Recordset
Valida_Cuenta = False
If Len(Trim(cta)) <> 7 And Trim(cta & "") <> "" Then
   MsgBox "Ingrese Cuenta Correctamente", vbInformation
   Exit Function
End If
If Trim(cta & "") <> "" Then
   If wGrupoPla = "01" And wcia = "21" Then
      Sql = "select cgcod from conmaspcge21 where cod_cia='" & wcia & "' and cgcod='" & cta & "'"
   Else
      Sql = "select cgcod from ." & wNomBd & "..conmaspcge where cod_cia='" & wcia & "' and cgcod='" & cta & "'"
   End If
   If Not fAbrRst(Rq, Sql) Then
      MsgBox "Cuenta no Registrada => " & cta, vbInformation
      Rq.Close: Set Rq = Nothing
      Exit Function
   End If
Rq.Close: Set Rq = Nothing
End If
Valida_Cuenta = True
End Function

Public Sub Carga_Utilidades(ano As Integer)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook


Dim xlApp1  As Object
Dim xlBook As Object

Sql$ = "usp_PlaCargaUtilidades '" & wcia & "'," & ano & ""
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub


Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("C:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 2).Value = NOMBREEMPRESA
xlSheet.Cells(3, 2).Value = "CALCULO DE PARTICIPACION DE UTILIDADES - " & Str(ano)
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 9)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 9)).HorizontalAlignment = xlCenter

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 9
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("D:I").ColumnWidth = 14
xlSheet.Range("D:I").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Range("j:j").NumberFormat = "m/d/yyyy"
xlSheet.Cells(5, 2).Value = "CODIGO"
xlSheet.Cells(5, 3).Value = "NOMBRE"
xlSheet.Cells(5, 4).Value = "TOT.REMUN."
xlSheet.Cells(5, 5).Value = "CALC. X REM."
xlSheet.Cells(5, 6).Value = "TOTAL HORAS"
xlSheet.Cells(5, 7).Value = "TOTAL DIAS"
xlSheet.Cells(5, 8).Value = "CALC. X DIAS"
xlSheet.Cells(5, 9).Value = "TOTAL"
xlSheet.Cells(5, 10).Value = "FECHA CESE"

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 10)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 10)).HorizontalAlignment = xlCenter
nFil = 6
Dim xItem As Integer
xItem = 1
Do While Not Rs.EOF
   xlSheet.Cells(nFil, 1).Value = xItem
   xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!PlaCod)
   xlSheet.Cells(nFil, 3).Value = "'" & Trim(Rs!nombre)
   xlSheet.Cells(nFil, 4).Value = Rs!totalremu
   xlSheet.Cells(nFil, 5).Value = Rs!calcporremu
   xlSheet.Cells(nFil, 6).Value = Rs!totalhoras
   xlSheet.Cells(nFil, 7).Value = Rs!totaldias
   xlSheet.Cells(nFil, 8).Value = Rs!calcpordias
   xlSheet.Cells(nFil, 9).Value = Rs!Total
   xlSheet.Cells(nFil, 10).Value = Rs!fechacese
   nFil = nFil + 1
   xItem = xItem + 1
   Rs.MoveNext
Loop

Dim msum As String
Dim xsum As Integer
xsum = 6
msum = (nFil - xsum) * -1
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 9).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 9)).Borders.LineStyle = xlContinuous

Rs.Close: Set Rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Screen.MousePointer = vbDefault
End Sub

Public Sub Carga_Recibos_Utilidades(ano As Integer, lDia As Integer, lMEs As Integer)

Dim nFil As Long
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook


Dim xlApp1  As Object
Dim xlBook As Object

Dim mfec
mfec = "Lima " & Format(lDia, "00") & " de " & Name_Month(Format(lMEs, "00")) & " de " & Trim(Str(ano + 1))

Sql$ = "usp_PlaCargaUtilidades '" & wcia & "'," & ano & ""
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Screen.MousePointer = vbHourglass

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("G:G").ColumnWidth = 14
xlSheet.Range("G:G").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

'xlSheet.Cells(3, 1).RowHeight = 10
'xlSheet.Cells(4, 1).RowHeight = 10
'xlSheet.Cells(11, 1).RowHeight = 10
'xlSheet.Cells(17, 1).RowHeight = 10
'xlSheet.Cells(29, 1).RowHeight = 10
'xlSheet.Cells(31, 1).RowHeight = 10
'xlSheet.Cells(38, 1).RowHeight = 10
'xlSheet.Cells(44, 1).RowHeight = 10

    Dim xRepLegal As String
    Dim rl As ADODB.Recordset
    Set rl = New ADODB.Recordset
    Sql$ = "select rep_nom from cia where cod_cia='" & wcia & "'"
    If (fAbrRst(rl, Sql$)) Then xRepLegal = Trim(rl!rep_nom & "")
    rl.Close: Set rl = Nothing

nFil = 1
Dim xTotDias As Double
Dim xTotRemun As Double
Dim xFilTot As Long
Dim xFilTot2 As Long
xTotDias = 0: xTotRemun = 0

Do While Not Rs.EOF
    DoEvents
  
   Debug.Print "reg " & CStr(nFil) & " de " & CStr(Rs.RecordCount)
   
   xlSheet.Cells(nFil, 1).Value = NOMBREEMPRESA
   xlSheet.Cells(nFil, 7).Value = "'" & Trim(Trim(Rs!fcese & "") & " ") & Trim(Rs!PlaCod)
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "Participación de Utilidades del Ejercicio " & Str(ano)
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "He recibido de la Compañía " & NOMBREEMPRESA & " la suma de"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
'   nFil = nFil + 1
'   xlSheet.Cells(nFil, 1).Value = "domiciliado en " & Trae_DIRECCION(wcia)
'   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
'   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
'
'   nFil = nFil + 1
'   xlSheet.Cells(nFil, 1).Value = "Debidamente representada por " & xRepLegal & " la suma de"
'   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
'   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
   
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "S/. " & Format(Rs!Total, "###,###,###.00") & " " & monto_palabras(Rs!Total) & " SOLES"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
   
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "por concepto de Participación de Utilidades del ejercicio " & Trim(Str(ano)) & ", calculo de acuerdo al"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
   
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Decreto Legislativo 892. Según como sigue:"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "Renta Anual a distribuir"
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!Monto, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Número de días laborados por el trabajador"
   xlSheet.Cells(nFil, 7).Value = Format(Rs!totaldias, "###,###,###")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Total Remuneración percibida por el trabajador ( " & Trim(Str(ano)) & " )"
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!totalremu, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xFilTot = nFil
   xlSheet.Cells(nFil, 1).Value = "Número total de días laborales por todos los trabajadores"
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   xlSheet.Cells(nFil, 7).Value = Format(Rs!totdias, "###,###,###")
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Remuneración total pagada a todos los trabajadores"
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 7).Value = Format(Rs!TotIng, "###,###,###.00")
   
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "1.- 50% Calculado de acuerdo a los días efectivos de trabajo."
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!calcpordias, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "2.- 50% Calculado de acuerdo al total de ingresos percibidos."
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!calcporremu, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Total a pagar"
   xlSheet.Cells(nFil, 1).Font.Bold = True
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 5)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!Total, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   xlSheet.Cells(nFil, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
   xlSheet.Cells(nFil, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
   nFil = nFil + 4
   xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
   xlSheet.Cells(nFil, 3).HorizontalAlignment = xlCenter
   xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Merge
   xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
   nFil = nFil + 1
   xlSheet.Cells(nFil, 6).Value = mfec
   
   nFil = nFil + 8
   
   xlSheet.Cells(nFil, 1).Value = NOMBREEMPRESA
   xlSheet.Cells(nFil, 7).Value = "'" & Trim(Trim(Rs!fcese & "") & " ") & Trim(Rs!PlaCod)
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "Participación de Utilidades del Ejercicio " & Str(ano)
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "He recibido de la Compañía " & NOMBREEMPRESA & " la suma de "
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
   nFil = nFil + 1
'   xlSheet.Cells(nFil, 1).Value = "domiciliado en " & Trae_DIRECCION(wcia)
'   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
'   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
'
'   nFil = nFil + 1
'   xlSheet.Cells(nFil, 1).Value = "Debidamente representada por " & xRepLegal & " la suma de"
'   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
'   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
'
'   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "S/. " & Format(Rs!Total, "###,###,###.00") & " " & monto_palabras(Rs!Total) & " SOLES"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
   
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "por concepto de Participación de Utilidades del ejercicio " & Trim(Str(ano)) & ", calculo de acuerdo al"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlDistributed
   
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Decreto Legislativo 892. Según como sigue:"
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Merge
   
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "Renta Anual a distribuir"
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!Monto, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Número de días laborados por el trabajador"
   xlSheet.Cells(nFil, 7).Value = Format(Rs!totaldias, "###,###,###")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Total Remuneración percibida por el trabajador ( " & Trim(Str(ano)) & " )"
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!totalremu, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xFilTot2 = nFil
   xlSheet.Cells(nFil, 1).Value = "Número total de días laborales por todos los trabajadores"
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   xlSheet.Cells(nFil, 7).Value = Format(Rs!totdias, "###,###,###")
   
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Remuneración total pagada a todos los trabajadores"
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 7).Value = Format(Rs!TotIng, "###,###,###.00")
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 2
   xlSheet.Cells(nFil, 1).Value = "1.- 50% Calculado de acuerdo a los días efectivos de trabajo."
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!calcpordias, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "2.- 50% Calculado de acuerdo al total de ingresos percibidos."
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!calcporremu, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   nFil = nFil + 1
   xlSheet.Cells(nFil, 1).Value = "Total a pagar"
   xlSheet.Cells(nFil, 1).Font.Bold = True
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 5)).Merge
   xlSheet.Cells(nFil, 1).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 6).Value = "S/."
   xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
   xlSheet.Cells(nFil, 7).Value = Format(Rs!Total, "###,###,###.00")
   xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
   xlSheet.Cells(nFil, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
   xlSheet.Cells(nFil, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
   nFil = nFil + 4
   xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
   xlSheet.Cells(nFil, 3).HorizontalAlignment = xlCenter
   xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Merge
   xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
   nFil = nFil + 1
   xlSheet.Cells(nFil, 6).Value = mfec
   
   nFil = nFil + 3
   xlSheet.HPageBreaks.Add (xlSheet.Cells(nFil - 2, 1))
   xTotDias = xTotDias + Rs!totaldias
   xTotRemun = xTotRemun + Rs!totalremu
   
   If nFil = 617 Then
        Stop
   End If
   
   Rs.MoveNext
Loop

xlSheet.Cells(xFilTot, 7).Value = Format(xTotDias, "###,###,###")
xlSheet.Cells(xFilTot + 1, 7).Value = Format(xTotRemun, "###,###,###")
xlSheet.Cells(xFilTot2, 7).Value = Format(xTotDias, "###,###,###")
xlSheet.Cells(xFilTot2 + 1, 7).Value = Format(xTotRemun, "###,###,###")
nFil = nFil + 1

Rs.Close: Set Rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

'If wGrupoPla = "01" And wcia = "21" Then cn21.Close

Screen.MousePointer = vbDefault
End Sub

Public Sub UtilidadesVsBoletas(ano As Integer)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object

Sql$ = "Usp_UtilidadesVsBoleta '" & wcia & "'," & ano & ""
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub


Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application

xlApp2.Sheets.Add
xlApp2.Sheets.Add
xlApp2.Sheets("Hoja1").Select
xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
xlApp2.Sheets("Hoja2").Select
xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)
xlApp2.Sheets("Hoja3").Select
xlApp2.Sheets("Hoja3").Move Before:=xlApp2.Sheets(3)


Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "TOTAL"

xlSheet.Range("C:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 2).Value = NOMBREEMPRESA
xlSheet.Cells(3, 2).Value = "UTILIDADES CORRESPONDIENTES AL PERIODO  " & Str(ano)
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 11)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 11)).HorizontalAlignment = xlCenter

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 9
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("D:I").ColumnWidth = 14
xlSheet.Range("D:J").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Range("K:K").NumberFormat = "m/d/yyyy"

xlSheet.Cells(5, 2).Value = "CODIGO"
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 2)).Merge

xlSheet.Cells(5, 3).Value = "NOMBRE"
xlSheet.Range(xlSheet.Cells(5, 3), xlSheet.Cells(6, 3)).Merge

xlSheet.Cells(5, 4).Value = "UTILIDADES"
xlSheet.Cells(6, 4).Value = "CALCULADAS"



xlSheet.Cells(5, 5).Value = "BOLETA DE PAGO POR UTILIDADES"
xlSheet.Range(xlSheet.Cells(5, 5), xlSheet.Cells(5, 11)).Merge



xlSheet.Cells(6, 5).Value = "UTILIDADES"
xlSheet.Cells(6, 6).Value = "TOTAL ING"
xlSheet.Cells(6, 7).Value = "CTA. CTE."
xlSheet.Cells(6, 8).Value = "RETENCION JUDICIAL"
xlSheet.Cells(6, 9).Value = "QUINTA CATEGORIA"
xlSheet.Cells(6, 10).Value = "NETO"
xlSheet.Cells(6, 11).Value = "F. CESE"

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 11)).WrapText = True

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 11)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Borders.LineStyle = xlNone
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 11)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 11)).VerticalAlignment = xlCenter

nFil = 7
Dim xItem As Integer
xItem = 1
Do While Not Rs.EOF
   xlSheet.Cells(nFil, 1).Value = xItem
   xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!PlaCod)
   xlSheet.Cells(nFil, 3).Value = "'" & Trim(Rs!nombre)
   xlSheet.Cells(nFil, 4).Value = Rs!Calc_Utilidad
   xlSheet.Cells(nFil, 5).Value = Rs!Bol_Utilidad
   xlSheet.Cells(nFil, 6).Value = Rs!Bol_TotIng
   xlSheet.Cells(nFil, 7).Value = Rs!Bol_Cta_Cte
   xlSheet.Cells(nFil, 8).Value = Rs!Bol_Ret_Jud
   xlSheet.Cells(nFil, 9).Value = Rs!Bol_Qta_Cat
   xlSheet.Cells(nFil, 10).Value = Rs!Bol_Neto
   xlSheet.Cells(nFil, 11).Value = Rs!fcese
   
   nFil = nFil + 1
   xItem = xItem + 1
   Rs.MoveNext
Loop

Dim msum As String
Dim xsum As Integer
xsum = 6
msum = (nFil - xsum) * -1
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 9).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 10).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 10)).Borders.LineStyle = xlContinuous


Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "BOLETAS"

xlSheet.Range("C:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 2).Value = NOMBREEMPRESA
xlSheet.Cells(3, 2).Value = "UTILIDADES CORRESPONDIENTES AL PERIODO  " & Str(ano)
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 10)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 10)).HorizontalAlignment = xlCenter

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 9
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("D:I").ColumnWidth = 14
xlSheet.Range("D:J").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

xlSheet.Cells(5, 2).Value = "CODIGO"
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 2)).Merge

xlSheet.Cells(5, 3).Value = "NOMBRE"
xlSheet.Range(xlSheet.Cells(5, 3), xlSheet.Cells(6, 3)).Merge

xlSheet.Cells(5, 4).Value = "UTILIDADES"
xlSheet.Cells(6, 4).Value = "CALCULADAS"



xlSheet.Cells(5, 5).Value = "BOLETA DE PAGO POR UTILIDADES"
xlSheet.Range(xlSheet.Cells(5, 5), xlSheet.Cells(5, 10)).Merge

xlSheet.Cells(6, 5).Value = "UTILIDADES"
xlSheet.Cells(6, 6).Value = "TOTAL ING"
xlSheet.Cells(6, 7).Value = "CTA. CTE."
xlSheet.Cells(6, 8).Value = "RETENCION JUDICIAL"
xlSheet.Cells(6, 9).Value = "QUINTA CATEGORIA"
xlSheet.Cells(6, 10).Value = "NETO"

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 10)).WrapText = True

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 10)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Borders.LineStyle = xlNone
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 10)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 10)).VerticalAlignment = xlCenter

nFil = 7
xItem = 1
If Rs.RecordCount Then Rs.MoveFirst
Do While Not Rs.EOF
   If Rs!Bol_TotIng <> 0 Then
        xlSheet.Cells(nFil, 1).Value = xItem
        xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!PlaCod)
        xlSheet.Cells(nFil, 3).Value = "'" & Trim(Rs!nombre)
        xlSheet.Cells(nFil, 4).Value = Rs!Calc_Utilidad
        xlSheet.Cells(nFil, 5).Value = Rs!Bol_Utilidad
        xlSheet.Cells(nFil, 6).Value = Rs!Bol_TotIng
        xlSheet.Cells(nFil, 7).Value = Rs!Bol_Cta_Cte
        xlSheet.Cells(nFil, 8).Value = Rs!Bol_Ret_Jud
        xlSheet.Cells(nFil, 9).Value = Rs!Bol_Qta_Cat
        xlSheet.Cells(nFil, 10).Value = Rs!Bol_Neto
        nFil = nFil + 1
        xItem = xItem + 1
   End If
   Rs.MoveNext
Loop


xsum = 6
msum = (nFil - xsum) * -1
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 9).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 10).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 10)).Borders.LineStyle = xlContinuous


Set xlSheet = xlApp2.Worksheets("HOJA3")
xlSheet.Name = "CESADOS"

xlSheet.Range("C:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 2).Value = NOMBREEMPRESA
xlSheet.Cells(3, 2).Value = "UTILIDADES CORRESPONDIENTES AL PERIODO  " & Str(ano)
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 4)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 4)).HorizontalAlignment = xlCenter

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 45

xlSheet.Range("D:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Range("E:E").NumberFormat = "m/d/yyyy"

xlSheet.Cells(5, 2).Value = "CODIGO"
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 2)).Merge

xlSheet.Cells(5, 3).Value = "NOMBRE"
xlSheet.Range(xlSheet.Cells(5, 3), xlSheet.Cells(6, 3)).Merge

xlSheet.Cells(5, 4).Value = "UTILIDADES CALCULADAS"
xlSheet.Range(xlSheet.Cells(5, 4), xlSheet.Cells(6, 4)).Merge

xlSheet.Cells(5, 5).Value = "F. CESE"
xlSheet.Range(xlSheet.Cells(5, 5), xlSheet.Cells(6, 5)).Merge

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 5)).WrapText = True

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 5)).Borders.LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 5)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(6, 5)).VerticalAlignment = xlCenter

nFil = 7
xItem = 1
If Rs.RecordCount Then Rs.MoveFirst
Do While Not Rs.EOF
   If Rs!Bol_TotIng = 0 Then
        xlSheet.Cells(nFil, 1).Value = xItem
        xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!PlaCod)
        xlSheet.Cells(nFil, 3).Value = "'" & Trim(Rs!nombre)
        xlSheet.Cells(nFil, 4).Value = Rs!Calc_Utilidad
        xlSheet.Cells(nFil, 5).Value = Rs!fcese
        nFil = nFil + 1
        xItem = xItem + 1
   End If
   Rs.MoveNext
Loop


xsum = 6
msum = (nFil - xsum) * -1
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 4)).Borders.LineStyle = xlContinuous


Rs.Close: Set Rs = Nothing

For I = 1 To 3
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridlines = False
   'xlApp2.ActiveWindow.Zoom = 80
Next
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Screen.MousePointer = vbDefault
End Sub

Public Sub Sustento_Utilidades(ano As Integer)
Dim nFil As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

Dim xlApp1  As Object
Dim xlBook As Object

Sql$ = "usp_PlaCalculaUtilidades_Sustento '" & wcia & "'," & ano & ""
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub


Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("C:D").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 2).Value = NOMBREEMPRESA
xlSheet.Cells(3, 2).Value = "SUSTENTO DE UTILIDADES - " & Str(ano)
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 5)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 5)).HorizontalAlignment = xlCenter

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:E").ColumnWidth = 15
xlSheet.Range("C:C").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Range("D:D").NumberFormat = "#,##0_ ;[Red]-#,##0 "
xlSheet.Range("E:E").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

xlSheet.Cells(5, 2).Value = "MES"
xlSheet.Cells(5, 3).Value = "HORAS"
xlSheet.Cells(5, 4).Value = "DIAS"
xlSheet.Cells(5, 5).Value = "INGRESOS"

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 5)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 5)).HorizontalAlignment = xlCenter
nFil = 8
Dim xItem As Integer
xItem = 1

Dim Rq As ADODB.Recordset
Dim mcod As String
mcod = Rs!PlaCod

nFil = 8

xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!PlaCod)
xlSheet.Cells(nFil, 2).Font.Bold = True
Sql = "select rtrim(ap_pat)+' '+rtrim(ap_mat)+' '+rtrim(nom_1)+' '+rtrim(nom_2) as Nombre From planillas where cia='" & wcia & "' and placod='" & Trim(Rs!PlaCod & "") & "' and status<>'*'"
If fAbrRst(Rq, Sql) Then xlSheet.Cells(nFil, 3).Value = Trim(Rq!nombre & "")
xlSheet.Cells(nFil, 3).Font.Bold = True
Rq.Close: Set Rq = Nothing
nFil = nFil + 2
Dim msum As String
Dim xsum As Integer
xsum = nFil

Dim lMEs As String
Do While Not Rs.EOF
   If mcod <> Rs!PlaCod Then
      msum = (nFil - xsum) * -1
      xlSheet.Cells(nFil, 3).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
      xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
      xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Borders.LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Font.Bold = True

      nFil = nFil + 2
      xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!PlaCod)
      xlSheet.Cells(nFil, 2).Font.Bold = True
      Sql = "select rtrim(ap_pat)+' '+rtrim(ap_mat)+' '+rtrim(nom_1)+' '+rtrim(nom_2) as Nombre From planillas where cia='" & wcia & "' and placod='" & Trim(Rs!PlaCod & "") & "' and status<>'*'"
      If fAbrRst(Rq, Sql) Then xlSheet.Cells(nFil, 3).Value = Trim(Rq!nombre & "")
      xlSheet.Cells(nFil, 3).Font.Bold = True
      Rq.Close: Set Rq = Nothing
      
      nFil = nFil + 2
      mcod = Rs!PlaCod
      xsum = nFil
   End If
   
   If Rs!Mes = 13 Then
      xlSheet.Cells(nFil, 2).Value = "Agregado"
   Else
      xlSheet.Cells(nFil, 2).Value = Name_Month(Format(Rs!Mes, "00"))
   End If
   
   xlSheet.Cells(nFil, 3).Value = Rs!horas
   xlSheet.Cells(nFil, 4).Value = Rs!Dias
   xlSheet.Cells(nFil, 5).Value = Rs!Ingreso
   xlSheet.Cells(nFil, 6).Value = "'" & Trim(Rs!PlaCod)
   nFil = nFil + 1
   xItem = xItem + 1
   Rs.MoveNext
Loop

msum = (nFil - xsum) * -1
xlSheet.Cells(nFil, 3).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 5)).Borders.LineStyle = xlContinuous


Rs.Close: Set Rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Screen.MousePointer = vbDefault
End Sub

