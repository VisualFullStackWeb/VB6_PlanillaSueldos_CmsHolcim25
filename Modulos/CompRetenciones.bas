Attribute VB_Name = "CompRetenciones"
Option Explicit
Dim s_CompRetenciones As String
Public rs_compRetenciones_Pub As ADODB.Recordset
Public Bol_ConDestino As Boolean

Sub Recupera_Jornal_Basico(CodCompañia As String, AñoProceso As String, CodTrabajador As String, _
Proceso As String)
    s_CompRetenciones = "Comprobante_Retenciones 1,'" & CodCompañia & "','" & AñoProceso & "'," & _
    "'" & CodTrabajador & "','" & Proceso & "'"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_IncrementoAfp(CodCompañia As String, AñoProceso As String, CodTrabajador As String)
    s_CompRetenciones = "Comprobante_Retenciones 4,'" & CodCompañia & "','" & AñoProceso & "'," & _
    "'" & CodTrabajador & "',''"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Gratificaciones(CodCompañia As String, AñoProceso As String, CodTrabajador As String, _
Proceso As String)
    s_CompRetenciones = "Comprobante_Retenciones 3,'" & CodCompañia & "','" & AñoProceso & "'," & _
    "'" & CodTrabajador & "','" & Proceso & "'"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Vacaciones(CodCompañia As String, AñoProceso As String, CodTrabajador As String, _
Proceso As String)
    s_CompRetenciones = "Comprobante_Retenciones 2,'" & CodCompañia & "','" & AñoProceso & "'," & _
    "'" & CodTrabajador & "','" & Proceso & "'"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Nombre_Trabajador(CodCompañia As String, CodTrabajador As String)
    s_CompRetenciones = "Comprobante_Retenciones 5,'" & CodCompañia & "','','" & CodTrabajador & "',''"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Informacion_Total(CodCompañia As String, AñoProceso As String, CodTrabajador As String)
    s_CompRetenciones = "Comprobante_Retenciones 6,'" & CodCompañia & "','" & AñoProceso & "'," & _
    "'" & CodTrabajador & "',''"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
'********codigo de conexion con la base de datos Starsoft*****************
Sub Conectar_Base_Datos_Access(CodEmpresa_Starsoft As String)
    '*************linea para ppm****************************
    cn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & Path_CIA & CodEmpresa_Starsoft & "\BDContabilidad.mdb"
    '*************linea para Roda***************************
    'cn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=I:\WENCO\DATA\" & CodEmpresa_Starsoft & "\BDContabilidad.mdb"
    cn2.Open
End Sub
Sub Cerrar_Conexion_Base_Datos_Access()
    cn2.Close
End Sub
Sub Recuperar_Cuentas_Naturaleza_Transferencia(CtaContable)
If Bol_ConDestino Then
    s_CompRetenciones = "select plancta_cargo1,plancta_abono1 from plan_cuenta_nacional " & _
    "where plancta_codigo='" & CtaContable & "'"
Else
    s_CompRetenciones = "select plancta_cargo1,plancta_abono1 from plan_cuenta_nacional " & _
    "where plancta_codigo='" & CtaContable & "99999999999'"
End If
    
    
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn2, adOpenKeyset, adLockOptimistic
End Sub



