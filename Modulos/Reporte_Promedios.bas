Attribute VB_Name = "Reporte_Promedios"
Option Explicit
Dim s_Promedios As String
Dim rs_Promedios As ADODB.Recordset
Public rs_Promedios_Pub As ADODB.Recordset
Sub Crea_Tabla_Promedios_Temporal(Año_Proceso As String, MesProceso_Par1 As Integer, MesProceso_Par2 As Integer, _
CodEmpresa As String, TipTrabajador As String, Concatenacion As String)
    s_Promedios = "delete from promedios_generales"
    cn.Execute s_Promedios
    
    If CodEmpresa = "06" Then
        s_Promedios = "insert into promedios_generales (cia,placod,nombres,fechaproceso,semana,i10,i11,i16,i24,i25) " & _
        "(select distinct p.cia,p.placod,(rtrim(a.ap_pat) +'  ' + rtrim(ap_mat) + '  ' + rtrim(nom_1)) " & _
        "as nombres,p.fechaproceso,p.semana," & Concatenacion & " from " & _
        "plahistorico p inner join planillas a on p.placod=a.placod where year(p.fechaproceso)='" & Año_Proceso & "' " & _
        "and month(p.fechaproceso)<= " & MesProceso_Par1 & " and month(p.fechaproceso)>" & MesProceso_Par2 & " and p.status <> '*' " & _
        "and p.cia='" & CodEmpresa & "' and p.tipotrab='" & TipTrabajador & "' and p.proceso='01')"
    Else
        s_Promedios = "insert into promedios_generales (cia,placod,nombres,fechaproceso,semana,i10,i11,i16,i24,i25,i21) " & _
        "(select distinct p.cia,p.placod,(rtrim(a.ap_pat) +'  ' + rtrim(ap_mat) + '  ' + rtrim(nom_1)) " & _
        "as nombres,p.fechaproceso,p.semana," & Concatenacion & " from " & _
        "plahistorico p inner join planillas a on p.placod=a.placod where year(p.fechaproceso)='" & Año_Proceso & "' " & _
        "and month(p.fechaproceso)< " & MesProceso_Par1 & " and month(p.fechaproceso)>=" & MesProceso_Par2 & " and p.status <> '*' " & _
        "and p.cia='" & CodEmpresa & "' and p.tipotrab='" & TipTrabajador & "' and p.proceso='01')"
    End If
    
    
    
    cn.Execute s_Promedios
End Sub
Sub Genera_Tabla_Promedios_Maestra1(Tipo_Trabajador As String, Cod_Empresa As String)
    cn.Execute "Reporte_Promedios 1,'" & Cod_Empresa & "','" & Tipo_Trabajador & "','02'," & _
    "'16','','','',0,0,0,0,0,0"
End Sub
Sub Recupera_Informacion_Maestra1()
    s_Promedios = "Reporte_Promedios 2,'','','','','','','',0,0,0,0,0,0"
    Set rs_Promedios_Pub = New ADODB.Recordset
    rs_Promedios_Pub.Open s_Promedios, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Informacion_Promedios_Generales()
    s_Promedios = "Reporte_Promedios 3,'','','','','','','',0,0,0,0,0,0"
    Set rs_Promedios_Pub = New ADODB.Recordset
    rs_Promedios_Pub.Open s_Promedios, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Graba_Informacion_Uniforme(Columna_Sumar As String, CodTrabajador As String, MesProceso As _
Integer)
    s_Promedios = "update promedios_generales set " & Columna_Sumar & "=(select sum(" & Columna_Sumar & ") from " & _
    "promedios_generales where placod='" & CodTrabajador & "' and month(fechaproceso)=" & MesProceso & ") " & _
    "where placod='" & CodTrabajador & "' and month(fechaproceso)=" & MesProceso & ""
    cn.Execute s_Promedios
End Sub
Sub Crear_Tabla_Reporte_Final()
    'cn.Execute "Reporte_Promedios 4,'','','','','','','',0,0,0,0,0,0"
End Sub
Function Verifica_Existencia_Registro_Promedios(CodCompañia As String, CodTrabajador As String, Descripcion _
As String) As Boolean
    Verifica_Existencia_Registro_Promedios = False
    s_Promedios = "Reporte_Promedios 5,'" & CodCompañia & "','','','','" & CodTrabajador & "'," & _
    "'" & Descripcion & "','',0,0,0,0,0,0"
    Set rs_Promedios = New ADODB.Recordset
    rs_Promedios.Open s_Promedios, cn, adOpenKeyset, adLockOptimistic
    If rs_Promedios.EOF = False Then
        Verifica_Existencia_Registro_Promedios = True
    End If
    Set rs_Promedios = Nothing
End Function
Sub Graba_Nueva_Informacion_Promedios(CodEmpresa As String, CodTrabajador As String, NombreTrab As String, _
Descripcion_Concepto As String, Monto1 As Single, Opcion As Integer)
    Select Case Opcion
        Case 1
            s_Promedios = "Reporte_Promedios 6,'" & CodEmpresa & "','','','','" & CodTrabajador & "'," & _
            "'" & Descripcion_Concepto & "','" & NombreTrab & "'," & Monto1 & ",0,0,0,0,0"
            cn.Execute s_Promedios
        Case 2
            s_Promedios = "Reporte_Promedios 6,'" & CodEmpresa & "','','','','" & CodTrabajador & "'," & _
            "'" & Descripcion_Concepto & "','" & NombreTrab & "',0," & Monto1 & ",0,0,0,0"
            cn.Execute s_Promedios
        Case 3
            s_Promedios = "Reporte_Promedios 6,'" & CodEmpresa & "','','','','" & CodTrabajador & "'," & _
            "'" & Descripcion_Concepto & "','" & NombreTrab & "',0,0," & Monto1 & ",0,0,0"
            cn.Execute s_Promedios
        Case 4
            s_Promedios = "Reporte_Promedios 6,'" & CodEmpresa & "','','','','" & CodTrabajador & "'," & _
            "'" & Descripcion_Concepto & "','" & NombreTrab & "',0,0,0," & Monto1 & ",0,0"
            cn.Execute s_Promedios
        Case 5
            s_Promedios = "Reporte_Promedios 6,'" & CodEmpresa & "','','','','" & CodTrabajador & "'," & _
            "'" & Descripcion_Concepto & "','" & NombreTrab & "',0,0,0,0," & Monto1 & ",0"
            cn.Execute s_Promedios
        Case 6
            s_Promedios = "Reporte_Promedios 6,'" & CodEmpresa & "','','','','" & CodTrabajador & "'," & _
            "'" & Descripcion_Concepto & "','" & NombreTrab & "',0,0,0,0,0," & Monto1 & ""
            cn.Execute s_Promedios
    End Select
End Sub
Sub Recupera_Promedios_Procesados_Por_Mes(CodEmpresa As String, MesProceso As String, Concatenado As String)
    s_Promedios = "select distinct cia,placod,nombres," & Concatenado & " from promedios_generales " & _
    "where month(fechaproceso)='" & MesProceso & "' and Cia='" & CodEmpresa & "'"
    Set rs_Promedios_Pub = New ADODB.Recordset
    rs_Promedios_Pub.Open s_Promedios, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Recupera_Monto_AFP_Diario(CodTrabajador As String, AñoProceso As String, MesProceso As String)
    s_Promedios = "select (i05 + i06)/30 as Monto from plahistorico where proceso='02' and cia='07' and status!='*' " & _
    "and placod='" & CodTrabajador & "' AND year(fecperiodovaca)=" & AñoProceso & " and month(fecperiodovaca)=" & _
    "" & MesProceso & ""
    Set rs_Promedios_Pub = New ADODB.Recordset
    rs_Promedios_Pub.Open s_Promedios, cn, adOpenKeyset, adLockOptimistic
End Sub
Sub Edita_Informacion_Promedios(CodEmpresa As String, CodTrabajador As String, NombreTrab As String, _
Descripcion_Concepto As String, Monto1 As Single, Opcion As Integer)
    Select Case Opcion
        Case 1
            s_Promedios = "update promedios_maestra2 set mes1=" & Monto1 & " where cia='" & CodEmpresa & "' " & _
            "and placod='" & CodTrabajador & "' and descripcion='" & Descripcion_Concepto & "'"
            cn.Execute s_Promedios
        Case 2
            s_Promedios = "update promedios_maestra2 set mes2=" & Monto1 & " where cia='" & CodEmpresa & "' " & _
            "and placod='" & CodTrabajador & "' and descripcion='" & Descripcion_Concepto & "'"
            cn.Execute s_Promedios
        Case 3
            s_Promedios = "update promedios_maestra2 set mes3=" & Monto1 & " where cia='" & CodEmpresa & "' " & _
            "and placod='" & CodTrabajador & "' and descripcion='" & Descripcion_Concepto & "'"
            cn.Execute s_Promedios
        Case 4
            s_Promedios = "update promedios_maestra2 set mes4=" & Monto1 & " where cia='" & CodEmpresa & "' " & _
            "and placod='" & CodTrabajador & "' and descripcion='" & Descripcion_Concepto & "'"
            cn.Execute s_Promedios
        Case 5
            s_Promedios = "update promedios_maestra2 set mes5=" & Monto1 & " where cia='" & CodEmpresa & "' " & _
            "and placod='" & CodTrabajador & "' and descripcion='" & Descripcion_Concepto & "'"
            cn.Execute s_Promedios
        Case 6
            s_Promedios = "update promedios_maestra2 set mes6=" & Monto1 & " where cia='" & CodEmpresa & "' " & _
            "and placod='" & CodTrabajador & "' and descripcion='" & Descripcion_Concepto & "'"
            cn.Execute s_Promedios
    End Select
End Sub
