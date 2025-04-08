Attribute VB_Name = "Reportes_Centrales"
Option Explicit
Dim s_RptCentrales As String
Dim rs_RptCentrales As ADODB.Recordset
Public rs_RptCentrales_pub As ADODB.Recordset
Public i_Direccion_Reportes As Integer
Sub Genera_Informacion_Reporte_Asistencia(CodCompañia As String, TipTrabajador As String, _
Dia As Integer, Numero_Pag As Integer)
     cn.Execute "Reportes 1,'" & CodCompañia & "','" & TipTrabajador & "','',''," & Numero_Pag & ",'','','','',''," & Dia & ""
End Sub
Sub Elimina_Informacion_Reporte_Asistencia()
    cn.Execute "Reportes 2,'','','','',0,'','','','','',0"
End Sub
Sub Ingresa_Nueva_Informacion(Nombre_Nuevo As String, Condicion_Busqueda As String, _
NumeroOrden As Integer, dias As Integer)
    cn.Execute "Reportes 3,'','','" & Nombre_Nuevo & "','" & Condicion_Busqueda & "'," & _
    "" & NumeroOrden & ",'','','','',''," & dias & ""
End Sub
Sub Recupera_Nombres_Reporte(dias As Integer)
    s_RptCentrales = "Reportes 4,'','','','',0,'','','','',''," & dias & ""
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Informacion_Empresa(CodEmpresa As String)
    s_RptCentrales = "Reportes 5,'" & CodEmpresa & "','','','',0,'','','','','',0"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Function Verifica_Existencia_Tabla_RegistroExistencia(CodEmpresa As String) As Boolean
    On Error GoTo Existencia
    Verifica_Existencia_Tabla_RegistroExistencia = True
    s_RptCentrales = "Reportes 7,'" & CodEmpresa & "','','','',0,'','','','','',0"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
    Exit Function
Existencia:
    Verifica_Existencia_Tabla_RegistroExistencia = False
End Function
Sub Crear_Tabla_RegistroAsistencia(CodCompañia As String)
    cn.Execute "Reportes 6,'" & CodCompañia & "','','','',0,'','','','','',0"
End Sub
Sub Inserta_Inforamcion_Nueva(CodCompañia As String, Horario As String, NumPagina As Integer)
    cn.Execute "Reportes 8,'" & CodCompañia & "','','',''," & NumPagina & ",'" & Horario & "'," & _
    "'','','','',0"
End Sub
Sub Recupera_Informacion_RegistroExistencia(CodEmpresa As String)
    s_RptCentrales = "Reportes 7,'" & CodEmpresa & "','','','',0,'','','','','',0"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Generar_Proceso_Tabla_Reporte_Lista_Trabajadores(CodEmpresa As String)
    cn.Execute "Reportes 9,'" & CodEmpresa & "','','','',0,'','','','','',0"
End Sub
Sub Eliminar_Tabla_Lista_Trabajadores()
    cn.Execute "Reportes 10,'','','','',0,'','','','','',0"
End Sub
Sub Crear_Plantilla_Reportes_Afp(CodEmpresa As String, MesProceso As String, AñoProceso As String)
    cn.Execute "Reporte_Aportes_Afp 1,'" & CodEmpresa & "','" & MesProceso & "'," & _
    "'" & AñoProceso & "',0,0,0,0,0,0,'',''"
End Sub
Sub Recupera_Descripcion_Administradoras_Pensiones()
    s_RptCentrales = "Reporte_Aportes_Afp 4,'','','',0,0,0,0,0,0,'',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recuperar_Monto_Total_Remuneracion_Asegurable(CodEmpresa As String, AñoProceso As String, _
MesProceso As String, CodAfp As String)
    s_RptCentrales = "Reporte_Aportes_Afp 5,'" & CodEmpresa & "','" & MesProceso & "'," & _
    "'" & AñoProceso & "',0,0,0,0,0,0,'" & CodAfp & "',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Codigo_Administradora_Pensiones(DescripcionAFP As String)
    s_RptCentrales = "Reporte_Aportes_Afp 6,'','','',0,0,0,0,0,0,'','" & DescripcionAFP & "'"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Enviar_Montos_Aportaciones_Asegurables(Rem_Asegurable As Double, Aporte_Obligatorio As Double, _
FondoPensiones As Double, PrimaSeguros As Double, ComisionRA As Double, Total_Retensiones As Double)
    cn.Execute "Reporte_Aportes_Afp 3,'','',''," & Rem_Asegurable & "," & Aporte_Obligatorio & "," & _
    "" & FondoPensiones & "," & PrimaSeguros & "," & ComisionRA & "," & Total_Retensiones & ",'',''"
End Sub
Sub Recupera_Factores_Operacionales(CodEmpresa As String, AñoProceso As String, MesProceso As _
String, CodAfp As String)
    s_RptCentrales = "Reporte_Aportes_Afp 7,'" & CodEmpresa & "','" & MesProceso & "'," & _
    "'" & AñoProceso & "',0,0,0,0,0,0,'" & CodAfp & "',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Elimina_Informacion_Reporte_Aportes()
    cn.Execute "Reporte_Aportes_Afp 2,'','','',0,0,0,0,0,0,'',''"
End Sub
Sub Recupera_Informacion_Calculo_Prima_Seguros(CodEmpresa As String, MesProceso As String, AñoProceso As String, _
CodAfp As String)
    s_RptCentrales = "Reporte_Aportes_Afp 8,'" & CodEmpresa & "','" & MesProceso & "'," & _
    "'" & AñoProceso & "',0,0,0,0,0,0,'" & CodAfp & "',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Ruc_Empresa_Reporte()
    s_RptCentrales = "Reporte_Aportes_Afp 9,'','','',0,0,0,0,0,0,'',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Function Verifica_Existencia_Reporte3()
    On Error GoTo pasalo
    s_RptCentrales = "Reporte_Aportes_Afp 9,'','','',0,0,0,0,0,0,'',''"
    Set rs_RptCentrales = New ADODB.Recordset
    rs_RptCentrales.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
    If rs_RptCentrales.EOF = False Then
        Call Elimina_Informacion_Reporte_Aportes
    End If
    Set rs_RptCentrales_pub = Nothing
    Exit Function
pasalo:
End Function
Sub Graba_Informacion_Nueva_Registro_Asistencia(CodEmpresa As String, Horario As String, _
Pagina As Integer)
    cn.Execute "Reportes 11,'" & CodEmpresa & "','','',''," & Pagina & ",'" & Horario & "','','','','',0"
End Sub
Sub Recupera_Nombre_Afp(CodAfp As String)
    s_RptCentrales = "Reporte_Aportes_Afp 10,'','','',0,0,0,0,0,0,'" & CodAfp & "',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Graba_Informacion_Asiento_Provision_Parametros(CodEmpresa As String, TipTrabajador As String, _
TipProvision As String, CtaProvision As String, CtaCosto As String)
    cn.Execute "Asientos_Provision 1,'" & CodEmpresa & "','" & TipTrabajador & "'," & _
    "'" & TipProvision & "','" & CtaProvision & "','" & CtaCosto & "','',''"
End Sub
Sub Edita_Informacion_Asiento_Provision_Parametros(CodEmpresa As String, TipTrabajador As String, _
TipProvision As String, CtaProvision As String, CtaCosto As String)
    cn.Execute "Asientos_Provision 3,'" & CodEmpresa & "','" & TipTrabajador & "'," & _
    "'" & TipProvision & "','" & CtaProvision & "','" & CtaCosto & "','',''"
End Sub
Function Verifica_Existencia_registro_Parametros(CodEmpresa As String, TipTrabajador As String, _
TipProvision As String) As Boolean
    Verifica_Existencia_registro_Parametros = True
    s_RptCentrales = "Asientos_Provision 2,'" & CodEmpresa & "','" & TipTrabajador & "'," & _
    "'" & TipProvision & "','','','',''"
    Set rs_RptCentrales = New ADODB.Recordset
    rs_RptCentrales.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
    If rs_RptCentrales.EOF = True Then
        Verifica_Existencia_registro_Parametros = False
    End If
    Set rs_RptCentrales = Nothing
End Function
Sub Recupera_Informacion_Parametros_Asiento_Provision(CodEmpresa As String, TipTrabajador As String, _
TipProvision As String)
    s_RptCentrales = "Asientos_Provision 2,'" & CodEmpresa & "','" & TipTrabajador & "'," & _
    "'" & TipProvision & "','','','',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Provision_Vacaciones(CodEmpresa As String, AñoProceso As String, MesProceso As String)
    s_RptCentrales = "Asientos_Provision 4,'" & CodEmpresa & "','','','','','" & AñoProceso & "'," & _
    "'" & MesProceso & "'"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Provision_Gratificaciones(CodEmpresa As String, AñoProceso As String, MesProceso As String)
    s_RptCentrales = "Asientos_Provision 6,'" & CodEmpresa & "','','','','','" & AñoProceso & "'," & _
    "'" & MesProceso & "'"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Provision_Liquidacion(CodEmpresa As String, AñoProceso As String, MesProceso As String)
    s_RptCentrales = "Asientos_Provision 5,'" & CodEmpresa & "','','','','','" & AñoProceso & "'," & _
    "'" & MesProceso & "'"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recuperar_Codigo_Empresa_Starsoft(CodEmpresa As String)
    s_RptCentrales = "Asientos_Provision 7,'" & CodEmpresa & "','','','','','',''"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub

Public Sub Trae_Cia_StarSoft(ByVal mCia As String, mYear As Integer)
    Dim Cadena As String
    Cadena = "SP_EMP_RELACIONADA '" & mCia & "'," & mYear & ""
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open Cadena, cn, adOpenKeyset, adLockOptimistic
End Sub

Public Function Cia_StarSoft(ByVal mCia As String, mYear As Integer) As String
    Dim Cadena As String
    Dim rs As ADODB.Recordset
    Cadena = "SP_EMP_RELACIONADA '" & mCia & "'," & mYear & ""
    Set rs = New ADODB.Recordset
    rs.Open Cadena, cn, adOpenKeyset, adLockOptimistic
    Cia_StarSoft = Trim(rs!EMP_ID)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Function

Sub Elimina_Registros_Existentes(CodEmpresa As String, TipBoleta As String)
    s_RptCentrales = "delete from asientos_pla where pla_cia='" & CodEmpresa & "' and pla_boleta='" & TipBoleta & "'"
    cn.Execute s_RptCentrales
End Sub
'******************codigo nuevo giovanni 15092007****************************************
Sub Recupera_Tipo_de_Tiempo(CodEmpresa As String, Cod_Trabajador As String)
    s_RptCentrales = "select a.*,b.flag2,b.descrip from plaremunbase a,maestros_2 b where a.cia='" & CodEmpresa & "' " & _
    "and a.tipo=b.cod_maestro2 and placod='" & Cod_Trabajador & "' and a.status<>'*' and concepto='02' " & _
    "and right(ciamaestro,3)= '076' ORDER BY cod_maestro2 "
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Dias_Trabajados(CodEmpresa As String, CodTrabajador As String, MesProceso _
As String, AñoProceso As String, TipoProceso As String, NumSemana As String, _
TipoTrabajador As String)
    s_RptCentrales = "fc_diastrabajados '" & CodEmpresa & "','" & CodTrabajador & "'," & _
    "'" & MesProceso & "','" & AñoProceso & " ','" & TipoProceso & "','" & NumSemana & "'," & _
    "'" & TipoTrabajador & "'"
    Set rs_RptCentrales_pub = New ADODB.Recordset
    rs_RptCentrales_pub.Open s_RptCentrales, cn, adOpenKeyset, adLockOptimistic
End Sub
'***************************************************************************************

