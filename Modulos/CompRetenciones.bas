Attribute VB_Name = "CompRetenciones"
Option Explicit
Dim s_CompRetenciones As String
Public rs_compRetenciones_Pub As ADODB.Recordset
Public Bol_ConDestino As Boolean

Sub Recupera_Jornal_Basico(CodCompa�ia As String, A�oProceso As String, CodTrabajador As String, _
Proceso As String)
    s_CompRetenciones = "Comprobante_Retenciones 1,'" & CodCompa�ia & "','" & A�oProceso & "'," & _
    "'" & CodTrabajador & "','" & Proceso & "'"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_IncrementoAfp(CodCompa�ia As String, A�oProceso As String, CodTrabajador As String)
    s_CompRetenciones = "Comprobante_Retenciones 4,'" & CodCompa�ia & "','" & A�oProceso & "'," & _
    "'" & CodTrabajador & "',''"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Gratificaciones(CodCompa�ia As String, A�oProceso As String, CodTrabajador As String, _
Proceso As String)
    s_CompRetenciones = "Comprobante_Retenciones 3,'" & CodCompa�ia & "','" & A�oProceso & "'," & _
    "'" & CodTrabajador & "','" & Proceso & "'"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Vacaciones(CodCompa�ia As String, A�oProceso As String, CodTrabajador As String, _
Proceso As String)
    s_CompRetenciones = "Comprobante_Retenciones 2,'" & CodCompa�ia & "','" & A�oProceso & "'," & _
    "'" & CodTrabajador & "','" & Proceso & "'"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Nombre_Trabajador(CodCompa�ia As String, CodTrabajador As String)
    s_CompRetenciones = "Comprobante_Retenciones 5,'" & CodCompa�ia & "','','" & CodTrabajador & "',''"
    Set rs_compRetenciones_Pub = New ADODB.Recordset
    rs_compRetenciones_Pub.Open s_CompRetenciones, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Informacion_Total(CodCompa�ia As String, A�oProceso As String, CodTrabajador As String)
    s_CompRetenciones = "Comprobante_Retenciones 6,'" & CodCompa�ia & "','" & A�oProceso & "'," & _
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



