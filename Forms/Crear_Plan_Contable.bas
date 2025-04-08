Attribute VB_Name = "Crear_Plan_Contable"
Option Explicit
Dim rs_PlanCont As ADODB.Recordset
Public rs_PlanCont_Pub As ADODB.Recordset
Public s_CodTipo_G As String
Public s_CtaContableNeto As String
Public s_CentroCostoPub As String
Public mCentroCosto As String
Dim s_PlanCont As String
Public i_Longitud_Centro_Costo As Integer
Sub Crea_Tablas_Contable_Maestras()
    '********crea las tablas maestras para contabilidad***********************
    cn.Execute "Mant_Contable_Maestras 2,'','','','','','',''"
    cn.Execute "Mant_Contable_Maestras 1,'','','','','','',''"
    cn.Execute "Mant_Maestras_Contables_Fijas 1,'','',''"
    cn.Execute "Generando_Asientos 3,'','',0,'','','','','','','',0,0,'','','','',''"
    Call Ingresar_Informacion_Fija
End Sub
Function Verifica_Existencia_Tablas() As Boolean
    On Error GoTo verifica
    Verifica_Existencia_Tablas = True
    s_PlanCont = "Mant_Contable_Maestras 5,'','','','','','',''"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenStatic, adLockReadOnly
    Set rs_PlanCont = Nothing
    Exit Function
verifica:
    Verifica_Existencia_Tablas = False
End Function
Sub Recupera_Conceptos_Ingresos(Codigo_Compañia As String, Opcion As Integer, Tipo_Trabajador As String)
    Select Case Opcion
        Case 1: s_PlanCont = "Recuperar_Informacion_Plan_Contable '" & Codigo_Compañia & "','02',1,'','',''"
        Case 2: s_PlanCont = "Recuperar_Informacion_Plan_Contable '" & Codigo_Compañia & "','',4,'I','" & Tipo_Trabajador & "',''"
    End Select
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Conceptos_Deduccion(Codigo_Compañia As String, Opcion As Integer, Tipo_Trabajador As _
String)
    Select Case Opcion
        Case 1: s_PlanCont = "Recuperar_Informacion_Plan_Contable '" & Codigo_Compañia & "','03',1,'','',''"
        Case 2: s_PlanCont = "Recuperar_Informacion_Plan_Contable '" & Codigo_Compañia & "','',4,'D','" & Tipo_Trabajador & "',''"
    End Select
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Conceptos_Aportes(Codigo_Compañia As String, Opcion As Integer, Tipo_Trabajador As _
String)
    Select Case Opcion
        Case 1: s_PlanCont = "Recuperar_Informacion_Plan_Contable '" & Codigo_Compañia & "','03',2,'','',''"
        Case 2: s_PlanCont = "Recuperar_Informacion_Plan_Contable '" & Codigo_Compañia & "','',4,'A','" & Tipo_Trabajador & "',''"
    End Select
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Graba_Informacion_Contable_Maestros1(Clave As String, concepto As String, Cod_Maestro As String)
    cn.Execute "Mant_Contable_Maestras 4,'" & Clave & "','" & concepto & "','','','','" & Cod_Maestro & "',''"
End Sub
Sub Graba_Informacion_Contable_Maestros2(Cod_Maestro As String, CentroCosto As String, CtaContable _
As String, CodCompañia As String)
    cn.Execute "Mant_Contable_Maestras 3,'','','" & CentroCosto & "','" & CtaContable & "','03'," & _
    "'" & Cod_Maestro & "','" & CodCompañia & "'"
End Sub
Sub Recupera_Tipos_Trabajadores()
    s_PlanCont = "Recuperar_Informacion_Plan_Contable '','',5,'','',''"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Codigo_Tipo_Trabajador(Descripcion_Tipo As String)
    s_PlanCont = "Recuperar_Informacion_Plan_Contable '','',6,'','','" & Descripcion_Tipo & "'"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
    s_CodTipo_G = rs_PlanCont!COD_MAESTRO2
    Set rs_PlanCont = Nothing
End Sub
Sub Ingresar_Informacion_Fija()
    Dim i_Contador As Integer: Dim s_clave As String: Dim s_Concepto As String
    Dim s_CtaContable As String
    For i_Contador = 1 To 6
         Select Case i_Contador
            Case 1: s_clave = "001": s_Concepto = "Cuenta Contable Normal": s_CtaContable = ""
            Case 2: s_clave = "002": s_Concepto = "Cuenta Contable Vacaciones": s_CtaContable = ""
            Case 3: s_clave = "003": s_Concepto = "Cuenta Contable Gratificacion": s_CtaContable = ""
            Case 4: s_clave = "004": s_Concepto = "Ipss Obreros": s_CtaContable = ""
            Case 5: s_clave = "005": s_Concepto = "Ipss Empleados": s_CtaContable = ""
            Case 6: s_clave = "006": s_Concepto = "Cuenta Contable CTS": s_CtaContable = ""
         End Select
         cn.Execute "Mant_Maestras_Contables_Fijas 2,'" & s_clave & "','" & s_Concepto & "'," & _
         "'" & s_CtaContable & "'"
    Next i_Contador
End Sub
Public Sub Grabar_Informacion_Nueva_Fijo(Cta_Contable As String, Clave As String, cia As String, Optional tipo As String = "")
    cn.Execute "Mant_Maestras_Contables_Fijas 3,'" & Clave & "','','" & Cta_Contable & "', '" & cia & "', '" & tipo & "'"
End Sub
Public Sub Recupera_Informacion_Contable_Fijo(Optional T_Trab As String = "")
    s_PlanCont = "Mant_Maestras_Contables_Fijas 4,'','','','" & wcia & "', '" & T_Trab & "'"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Codigos_Conceptos_Ingresos(TipoOperacion As String, TipoTrabajador As String, CodEmpresa As _
String)
    s_PlanCont = "Generando_Asientos 1,'" & TipoOperacion & "','" & TipoTrabajador & "',0,'',''," & _
    "'','" & CodEmpresa & "','','','',0,0,'','','','','',''"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Graba_Informacion_Totales(Clave_Concepto As String, Año_Proceso As String, Mes_Proceso As _
String, Proceso As String, TipoTrabajador As String, Cod_Trabajador As String, SemProceso As String, _
CodEmpresa As String)
    s_PlanCont = "update contable_maestras2 set total_montos=(select sum(" & Clave_Concepto & ") " & _
    "as Monto from plahistorico where year(fechaproceso)=" & Año_Proceso & " and month" & _
    "(fechaproceso)=" & Mes_Proceso & " and proceso='" & Proceso & "' and placod='" & Cod_Trabajador & "' " & _
    "and status<>'*' and semana='" & SemProceso & "' and cia='" & CodEmpresa & "' ) where cma_codmaestro2='" & Clave_Concepto & TipoTrabajador & "' " & _
    "and cma_codcia='" & CodEmpresa & "' "
    cn.Execute s_PlanCont
End Sub

Sub Graba_Informacion_Total(Clave_Concepto As String, FechaProceso As _
Date, Proceso As String, TipoTrabajador As String, Cod_Trabajador As String, SemProceso As String, _
CodEmpresa As String)
    s_PlanCont = "SET DATEFORMAT DMY " & _
    "update contable_maestras2 set total_montos=(select sum(" & Clave_Concepto & ") " & _
    "as Monto from plahistorico where fechaproceso='" & Format(FechaProceso, "dd/MM/yyyy") & "' " & _
    " and proceso='" & Proceso & "' and placod='" & Cod_Trabajador & "' " & _
    "and status<>'*' and semana='" & SemProceso & "' and cia='" & CodEmpresa & "' ) where cma_codmaestro2='" & Clave_Concepto & TipoTrabajador & "' " & _
    "and cma_codcia='" & CodEmpresa & "' "
    cn.Execute s_PlanCont
End Sub

Sub Recupera_Informacion_Plan_Contable(Cta_Contable As String)
    's_PlanCont = "Generando_Asientos 5,'','',0,'','" & Cta_Contable & "','','','','',''," & _
    '"0,0,'','','','',''"
    'Set rs_PlanCont_Pub = New ADODB.Recordset
    'rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Function Verifica_Existencia_Registro(CodEmpresa As String, CodTrabajador As String, Año_Proceso _
As String, Mes_Proceso As String, Semana_Proceso As String, Voucher As String, _
Cta_Contable As String, TipBoleta As String, CentroCosto As String) As Boolean
    Verifica_Existencia_Registro = False
    s_PlanCont = "Generando_Asientos 6,'','',0,'','" & Cta_Contable & "','" & Mes_Proceso & "'," & _
    "'" & CodEmpresa & "','" & Voucher & "','','',0,0,'" & Año_Proceso & "','" & Semana_Proceso & "'," & _
    "'" & CodTrabajador & "','','" & TipBoleta & "','" & CentroCosto & "'"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
    If rs_PlanCont.EOF = False Then
        Verifica_Existencia_Registro = True
    End If
    Set rs_PlanCont = Nothing
End Function
Sub Graba_Asiento_Contable(CodEmpresa As String, CodTrabajador As String, Año_Proceso As String, _
Mes_Proceso As String, Semana_Proceso As String, Voucher As String, Cta_Contable As String, _
TipoAsiento As String, Descripcion_Cta As String, Haber_MN As Double, Debe_MN As Double, TipBoleta _
As String, TipoTrabajador As String, CentroCosto As String)
    cn.Execute "Generando_Asientos 7,'','" & TipoTrabajador & "',0,'','" & Cta_Contable & "'," & _
    "'" & Mes_Proceso & "','" & CodEmpresa & "','" & Voucher & "','" & Descripcion_Cta & "'," & _
    "'" & TipoAsiento & "'," & Haber_MN & "," & Debe_MN & ",'" & Año_Proceso & "'," & _
    "'" & Semana_Proceso & "','" & CodTrabajador & "','','" & TipBoleta & "','" & CentroCosto & "'"
End Sub
Sub Editar_Asiento_Contable(Monto As Double, CodEmpresa As String, CodTrabajador As String, _
Año_Proceso As String, Mes_Proceso As String, Semana_Proceso As String, Voucher As String, _
Cta_Contable As String, TipoBoleta As String, CentroCosto As String)
', Tipo As Integer
    'Select Case Tipo
     '   Case 1
            cn.Execute "Generando_Asientos 8,'',''," & Monto & ",'','" & Cta_Contable & "','" & Mes_Proceso & "'," & _
            "'" & CodEmpresa & "','" & Voucher & "','','',0,0,'" & Año_Proceso & "','" & Semana_Proceso & "'," & _
            "'" & CodTrabajador & "','','" & TipoBoleta & "','" & CentroCosto & "'"
      '  Case 2
       '     cn.Execute "Generando_Asientos 14,'',''," & Monto & ",'','" & Cta_Contable & "','" & Mes_Proceso & "'," & _
        '    "'" & CodEmpresa & "','" & Voucher & "','','',0,0,'" & Año_Proceso & "','" & Semana_Proceso & "'," & _
         '   "'" & CodTrabajador & "','',''"
    'End Select
End Sub

Sub Editar_Asiento_Contable2(Monto As Double, CodEmpresa As String, CodTrabajador As String, _
Año_Proceso As String, Mes_Proceso As String, Semana_Proceso As String, Voucher As String, _
Cta_Contable As String, TipBoleta As String, CentroCosto As String)
'            cn.Execute "Generando_Asientos 8,'',''," & Monto & ",'','" & Cta_Contable & "','" & Mes_Proceso & "'," & _
 '           "'" & CodEmpresa & "','" & Voucher & "','','',0,0,'" & Año_Proceso & "','" & Semana_Proceso & "'," & _
  '          "'" & CodTrabajador & "','',''"
    cn.Execute "Generando_Asientos 14,'',''," & Monto & ",'','" & Cta_Contable & "','" & Mes_Proceso & "'," & _
    "'" & CodEmpresa & "','" & Voucher & "','','',0,0,'" & Año_Proceso & "','" & Semana_Proceso & "'," & _
    "'" & CodTrabajador & "','','" & TipBoleta & "','" & CentroCosto & "'"
End Sub

Sub Recupera_Asiento_Contable_Neto(CodAsientoFijo As String, T_Trab As String)
    s_PlanCont = "Generando_Asientos 2,'','" & T_Trab & "',0,'" & CodAsientoFijo & "','','','" & wcia & "','','','',0,0," & _
    "'','','','','',''"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
    s_CtaContableNeto = rs_PlanCont!cmaf_ctacontable
    Set rs_PlanCont = Nothing
End Sub

Sub Recupera_Monto_Neto(Año_Proceso As String, Mes_Proceso As String, Tipo_Proceso As String, _
CodTrabajador As String, SemProceso As String, FechaProceso As Date)
    s_PlanCont = "Generando_Asientos 4,'','',0,'','','" & Mes_Proceso & "'," & _
    "'','','','',0,0,'" & Año_Proceso & "','" & SemProceso & "','" & CodTrabajador & "'," & _
    "'" & Tipo_Proceso & "','','','" & FechaProceso & "'"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub

Sub Recupera_Informacion_Total_Boletas(Año_Proceso As String, Mes_Proceso As String, TipoProceso _
As String, CodEmpresa As String, TipoTrabajador As String)
    s_PlanCont = "Generando_Asientos 9,'','" & TipoTrabajador & "',0,'','','" & Mes_Proceso & "','" & CodEmpresa & "'," & _
    "'','','',0,0,'" & Año_Proceso & "','','','" & TipoProceso & "','',''"
    's_PlanCont = "select * from plahistorico where year(fechaproceso)='2007' and month(fechaproceso)='07' " & _
    '"and proceso='01' and status <> '*' and cia='06' and tipotrab='01' and placod in ('RE242','RE243') order by placod,semana"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
Function Verifica_Ejecucion_Proceso(CodEmpresa As String, TipTrabajador As String, AñoProceso _
As String, MesProceso As String, TipBoleta As String) As Boolean
    Verifica_Ejecucion_Proceso = False
    s_PlanCont = "Generando_Asientos 10,'','" & TipTrabajador & "',0,'','','" & MesProceso & "'," & _
    "'" & CodEmpresa & "','','','',0,0,'" & AñoProceso & "','','','','" & TipBoleta & "',''"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
    If rs_PlanCont.EOF = False Then
        Verifica_Ejecucion_Proceso = True
    End If
    Set rs_PlanCont = Nothing
End Function
Sub Elimina_Informacion_Proceso_Anterior(CodEmpresa As String, TipTrabajador As String, AñoProceso _
As String, MesProceso As String, TipBoleta As String)
    cn.Execute "Generando_Asientos 11,'','" & TipTrabajador & "',0,'','','" & MesProceso & "'," & _
    "'" & CodEmpresa & "','','','',0,0,'" & AñoProceso & "','','','','" & TipBoleta & "',''"
End Sub
Function Verifica_y_Captura_Longitud_Centro_Costo(CodCompañia As String) As Boolean
    '*********se debe verificar como se llena la informacion del maestro y del maestro_2***********
    Verifica_y_Captura_Longitud_Centro_Costo = True
    s_PlanCont = "Cuenta_Contable_Centro_Costo 1,'" & CodCompañia & "'"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
    If rs_PlanCont.EOF = False Then
        If rs_PlanCont.RecordCount > 1 Then
            Verifica_y_Captura_Longitud_Centro_Costo = False
        Else
            rs_PlanCont.MoveFirst
            i_Longitud_Centro_Costo = rs_PlanCont!cantidad
        End If
    Else
        Verifica_y_Captura_Longitud_Centro_Costo = False
    End If
    Set rs_PlanCont = Nothing
End Function
Function Recupera_Centro_Costo(CodEmpresa As String, TipProceso As String) As Boolean
    On Error GoTo centro
    Recupera_Centro_Costo = True
    s_PlanCont = "Generando_Asientos 12,'','',0,'','','','" & CodEmpresa & "','','','',0,0,'',''," & _
    "'','" & TipProceso & "','',''"
    Set rs_PlanCont = New ADODB.Recordset
    rs_PlanCont.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
    If rs_PlanCont.RecordCount > 0 Then
        s_CentroCostoPub = rs_PlanCont!flag1
        mCentroCosto = rs_PlanCont!CENCOSTO
    Else
        Recupera_Centro_Costo = False
    End If
    Set rs_PlanCont = Nothing
    Exit Function
centro:
    MsgBox ERR.Description & ", Ocurrio un Error con el Centro de Costo, Codigo Interno del Centro de Costo no Existe", vbCritical
    Recupera_Centro_Costo = False
End Function

Sub Recupera_Informacion_ImportacionDBF(CodEmpresa As String, AñoProceso As String, MesProceso As _
String)
    s_PlanCont = "Generando_Asientos 13,'','',0,'','','" & MesProceso & "','" & CodEmpresa & "','','','',0,0," & _
    "'" & AñoProceso & "','','','','',''"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub

Sub Recupera_Fecha_Proceso()
    s_PlanCont = "select distinct fechaproceso from plahistorico where year(fechaproceso)='2007' " & _
    "and month(fechaproceso)='07' and day(fechaproceso)='31'"
    Set rs_PlanCont_Pub = New ADODB.Recordset
    rs_PlanCont_Pub.Open s_PlanCont, cn, adOpenKeyset, adLockOptimistic
End Sub
