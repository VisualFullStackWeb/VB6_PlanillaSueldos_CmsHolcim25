VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCalculaBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Sql                 As String
Private m_NuevaBoleta       As Boolean
Private m_RsBoleta          As New ADODB.Recordset
Private m_CodTrabajor       As String
Private m_CodAuxiInterno    As String
Private m_TipoTrabajor      As String
Private m_TipoBoleta        As String
Private m_FechaProceso      As String
Private m_FechaInicio_Vacaciones As String
Private m_FechaFin_Vacaciones As String
Private m_NroSemana         As Integer
Private m_Turno             As String
Private m_CodCia            As String
Private m_PagoQuincena      As Boolean 'WTIPDOC
Private m_Status            As String
Private m_CodAfp            As String
Private m_NumAfp                    As String
Private m_Obra                      As String
Private m_ImpBasico                 As Currency
Private m_FechaPeriodo_Vacaciones   As String
Private m_Fecha_Ingreso             As String
Private m_Fecha_Nacimiento          As String
Private m_año_proceso               As Integer
Private m_mes_proceso               As Integer
Private m_dia_proceso               As Integer

Private m_TopeHrsExtras             As String
Private m_CargoTrab                 As String
Private m_Vacacion                  As String
Private m_CancelaGrabacion          As Boolean
Private m_VokDevengue               As Boolean
Private m_Fecha_Jubilacion          As String
Private m_Sn_Essaludvida            As Byte

Private m_Sn_sindicato              As Byte
Private m_Manos                     As Long
Private m_Sn_Quinta                 As Boolean
Private m_Horas                     As Integer
Private m_CodArea                   As String
Private m_CalcularH14               As Boolean
Private m_TipoPeriodoPago           As String
Private m_RecalculoBoleta           As Boolean
Private m_Calcular_Accidente_Trabajo   As Boolean
Private XB_ASIGFAM As Boolean
Dim ArrDsctoCTACTE() As Variant
' Api GetSafeArrayPointer para saber si se incializo el array
Private Declare Sub GetSafeArrayPointer Lib "msvbvm60.dll" Alias "GetMem4" (pArray() As Any, ret As Long)

'/**********************************************/
'/*** FUNCION PRINCIPAL DE CALCULO DE BOLETA ***/
'/**********************************************/

Public Function CalcularBoleta(ByVal pm_CodCia As String, ByVal m_NuevaBoleta As Boolean _
, ByRef pRsHoras As ADODB.Recordset, ByRef pRsCosto As ADODB.Recordset, ByRef pRsPagAdic As ADODB.Recordset, ByRef pRsDesAdic As ADODB.Recordset _
, ByVal pm_CodTrabajor As String, ByVal pm_TipoTrabajor As String, ByVal pm_TipoBoleta As String _
, ByVal pm_FechaProceso As String, ByVal pm_FechaInicio_Vacaciones As String, ByVal pm_FechaFin_Vacaciones As String _
, ByVal pm_NroSemana As Integer, ByVal pm_Turno As String, ByVal pm_PagoQuincena As Boolean _
, ByVal pm_Status As String, ByVal pm_CodObra As String, ByVal pm_FechaPeriodo_Vacaciones As String _
, ByVal pm_CalcularH14 As Boolean, ByVal pm_RecalculoBoleta As Boolean _
, ByRef pm_ArrDsctoCtaCte() As Variant, _
  Optional CargaPlanilla As Boolean = False, _
  Optional B_ASIGFAM As Boolean = False, _
  Optional pNumQuincena As String = "") As Boolean

XB_ASIGFAM = B_ASIGFAM

'/*** Control de Error en los procedimientos y funciones ***/

Dim SwError As Boolean
SwError = False

If pRsHoras Is Nothing Then Crea_Rs pRsHoras, pRsCosto, pRsPagAdic, pRsDesAdic


'/****** Inicializa Variables Externas
m_CodTrabajor = pm_CodTrabajor
m_TipoTrabajor = pm_TipoTrabajor
m_TipoBoleta = pm_TipoBoleta '01=Normal,02=Vacaciones,03=Gratificacion
m_FechaProceso = pm_FechaProceso
m_FechaInicio_Vacaciones = pm_FechaInicio_Vacaciones
m_FechaFin_Vacaciones = pm_FechaFin_Vacaciones
m_NroSemana = pm_NroSemana
m_Turno = pm_Turno
m_CodCia = pm_CodCia
m_PagoQuincena = pm_PagoQuincena
m_Status = pm_Status
m_Obra = pm_CodObra
m_RecalculoBoleta = pm_RecalculoBoleta
Dim s5taCargada As Boolean

s5taCargada = False

If Trim(m_CodCia & "") = "" Then
    MsgBox "Ingrese Codigo de Compañía correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
End If

If Trim(m_TipoTrabajor & "") = "" Then
    MsgBox "Ingrese Tipo de Trabajador correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
End If
If Trim(m_TipoBoleta & "") = "" Then
    MsgBox "Ingrese Tipo de Boleta correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
End If

If Not IsDate(m_FechaProceso & "") Then
    MsgBox "Ingrese Fecha de Proceso de la Boleta correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
End If

If Trim(m_TipoBoleta & "") = "02" Then 'vacaciones
    If Not IsDate(m_FechaInicio_Vacaciones & "") Then
        MsgBox "Ingrese Fecha Inicio de Vacaciones de la Boleta correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
    End If
    If Not IsDate(m_FechaFin_Vacaciones & "") Then
        MsgBox "Ingrese Fecha Fin de Vacaciones de la Boleta correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
    End If
End If

If Not IsNumeric(m_NroSemana & "") Then
    MsgBox "Ingrese Nro de Semana de la Boleta correctamente", vbExclamation, "Validación de Datos": GoTo ErrCalc:
End If

m_ImpBasico = Traer_SueldoBasico_Trabajador(m_CodCia, m_CodTrabajor, "01", SwError)
If SwError Then GoTo ErrCalc:

m_FechaPeriodo_Vacaciones = pm_FechaPeriodo_Vacaciones
m_CalcularH14 = pm_CalcularH14

'/*** Datos Generales del TRbajador ***//
Dim xciamae As String
xciamae = Traer_TipoTablaMaestra(m_CodCia, "01055")
Sql = Funciones.nombre()
Sql = Sql & " codauxinterno,a.status,a.tipotrabajador,a.fingreso,"
Sql = Sql & " a.fcese,a.codafp,a.numafp,a.area,a.placod,a.codauxinterno,b.descrip,a.tipotasaextra,"
Sql = Sql & " a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento,"
Sql = Sql & " a.fec_jubila,a.sindicato,a.ESSALUDVIDA,a.quinta,a.trab_Calc_AccidenteTrabajo, a.codsctr "
Sql = Sql & " from planillas a,maestros_2 b where a.status<>'*' "
Sql = Sql & xciamae
Sql = Sql & " and a.tipotrabajador=b.cod_maestro2 "
Sql = Sql & " and cia='" & m_CodCia & "' AND placod='" & Trim(m_CodTrabajor) & "' "
Sql = Sql & " order by nombre"

Dim Rt As ADODB.Recordset

If fAbrRst(Rt, Sql) Then
      'cargamos si la person esta afecto o no a la quinta categoria
      m_Sn_Quinta = True
      If Trim(Rt!quinta & "") = "N" Or Trim(Rt!quinta & "") = "" Then m_Sn_Quinta = False
      '***********************************************************
      
      m_Sn_Essaludvida = 1
      If Trim(Rt!essaludvida & "") = "N" Or Trim(Rt!essaludvida & "") = "" Then m_Sn_Essaludvida = 0
      
      m_Sn_sindicato = 1 '--ojo revisar este flag
      If Trim(Rt!sindicato & "") = "N" Or Trim(Rt!sindicato & "") = "" Then m_Sn_sindicato = 0
      
      m_CodAuxiInterno = Trim(Rt!codauxinterno & "")
      m_CodAfp = Trim(Rt!CodAfp & "")
      m_NumAfp = Trim(Rt!NUMAFP & "")
      m_Fecha_Nacimiento = Format(Rt!fnacimiento, "dd/mm/yyyy")
      m_Fecha_Ingreso = Format(Rt!fingreso, "dd/mm/yyyy")
      If Not IsNull(Rt!fec_jubila) Then
        m_Fecha_Jubilacion = Format(Rt!fec_jubila, "dd/mm/yyyy")
      Else
        m_Fecha_Jubilacion = ""
      End If
      m_TopeHrsExtras = Trim(Rt!tipotasaextra & "")
      m_CargoTrab = Trim(Rt!cargo & "")
      m_Vacacion = Trim(Rt!vacacion & "") 'prorrateo de vacaciones
      m_CodArea = Trim(Rt!Area & "")
      
      m_Manos = Cantidad_ymd_Trasncurridos(m_FechaProceso, m_Fecha_Nacimiento, "a")
      
      'DETERMINAR SI CALCULA SCTR
      
      m_Calcular_Accidente_Trabajo = IIf(Trim(Rt!codsctr & "") <> "", 1, 0)
      
      'm_Calcular_Accidente_Trabajo = IIf(Trim(Rt!trab_Calc_AccidenteTrabajo & "") = True, 1, 0)
      
      
      
Else
    MsgBox "No existen Datos generales del Trabajador " & Trim(m_CodTrabajor), vbCritical, "Datos Generales"
    GoTo ErrCalc:
End If

xciamae = Traer_TipoTablaMaestra(m_CodCia, "01055")

Sql = "Select flag1,flag2 from maestros_2 where cod_maestro2='" & m_TipoTrabajor & "' and status<>'*'"
Sql = Sql & xciamae

If fAbrRst(Rt, Sql) Then
    m_Horas = Val(Rt!flag2)
    m_TipoPeriodoPago = Left(Rt!flag1, 2)
Else
    MsgBox "No existen Nro de Horas por Tipo de Trabajador " & Trim(m_CodTrabajor), vbCritical, "Datos Generales"
    GoTo ErrCalc:
End If

Rt.Close
Set Rt = Nothing
           
m_año_proceso = Val(Mid(m_FechaProceso, 7, 4))
m_mes_proceso = Val(Mid(m_FechaProceso, 4, 2))
m_dia_proceso = Val(Mid(m_FechaProceso, 1, 2))

m_CancelaGrabacion = False

Screen.MousePointer = 11
cn.Execute "begin transaction", 64

If m_NuevaBoleta = False Then
     
     If m_RecalculoBoleta Then
        If Not CargaDatos_RecalculoBoleta(m_CodCia, m_CodTrabajor, m_TipoBoleta, m_NroSemana, m_año_proceso, m_mes_proceso, m_PagoQuincena, m_TipoPeriodoPago, pRsHoras, pRsCosto, pRsPagAdic, pRsDesAdic) = False Then GoTo ErrCalc:
    End If
    
    If Not EliminaBoleta(m_CodCia, m_PagoQuincena, m_año_proceso, m_mes_proceso, m_dia_proceso, Trim(m_CodTrabajor), m_TipoPeriodoPago, m_TipoBoleta, m_NroSemana, m_RecalculoBoleta, pNumQuincena) Then GoTo ErrCalc:
    
    If pm_RecalculoBoleta Then
        'If Not Calcular_Dscto_CtaCte_Boleta(pRsDesAdic, m_CodCia, m_CodTrabajor, m_TipoBoleta, m_FechaProceso, m_PagoQuincena) Then GoTo ErrCalc:
    Else
        ArrDsctoCTACTE = pm_ArrDsctoCtaCte
    End If
    
Else
    ArrDsctoCTACTE = pm_ArrDsctoCtaCte
End If
 
        
Sql$ = "delete from platemphist where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "'"
Sql$ = Sql$ & " and codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & Trim(m_TipoBoleta) & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(Trim(m_FechaProceso), FormatFecha) & "' and semana='" & Trim(m_NroSemana) & "' and status='" & Trim(m_Status) & "'"
cn.Execute Sql
        
        

Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql = Sql & "insert into platemphist(cia,placod,"
Sql = Sql & " codauxinterno,proceso,fechaproceso,fechavacai,fechavacaf,fecperiodovaca,semana,"
Sql = Sql & " fechaingreso,turno,codafp,status,fec_crea,"
Sql = Sql & " tipotrab,obra,numafp,basico,fec_modi,"
Sql = Sql & " user_crea,user_modi) "
Sql = Sql & " values('" & m_CodCia & "','" & Trim(m_CodTrabajor) & "','"
Sql = Sql & Trim(m_CodAuxiInterno) & "','" & Trim(m_TipoBoleta) & "','"
Sql = Sql & Format(Trim(m_FechaProceso), FormatFecha) & "'," & IIf(Trim(m_FechaInicio_Vacaciones) <> "__/__/____" And Trim(m_FechaInicio_Vacaciones) <> "", "'" & Format(Trim(m_FechaInicio_Vacaciones), FormatFecha) & "'", "NULL")
Sql = Sql & "," & IIf(Trim(m_FechaFin_Vacaciones) <> "__/__/____" And Trim(m_FechaFin_Vacaciones) <> "", "'" & Format(Trim(m_FechaFin_Vacaciones), FormatFecha) & "'", "NULL") & ","
Sql = Sql & IIf(Trim(m_FechaPeriodo_Vacaciones) <> "__/__/____" And Trim(m_FechaPeriodo_Vacaciones) <> "", "'" & Format(Trim(m_FechaPeriodo_Vacaciones), FormatFecha) & "'", "NULL")
Sql = Sql & ",'" & m_NroSemana & "','" & Format(m_Fecha_Ingreso, FormatFecha)
Sql = Sql & "','" & Format(m_Turno, "0") & "','"
Sql = Sql & m_CodAfp & "','" & m_Status & "',"
Sql = Sql & "getdate(),'" & m_TipoTrabajor & "','" & m_Obra
Sql = Sql & "','" & Trim(m_NumAfp) & "',"
Sql = Sql & CCur(m_ImpBasico) & ",getdate()"
Sql = Sql & ",'" & wuser & "','" & wuser & "')"
cn.Execute Sql

'Factor necesario para carga de quincena

Dim intFactorCarga As Integer
'intFactorCarga = IIf(wtipodoc = True, 1, IIf(pNumQuincena = "1", 2, 1))

Dim MqueryH As String
'Horas
MqueryH = ""
With pRsHoras
    If .RecordCount > 0 Then
       .MoveFirst
       Do While Not .EOF
            MqueryH = MqueryH & "h" & !codigo & "=" & IIf(IsNull(!Monto), 0, !Monto) & ""
          .MoveNext
          If Not .EOF Then MqueryH = MqueryH & ","
       Loop
    End If
End With

'Pagos Adicionales
Dim MqueryP As String
MqueryP = ""
With pRsPagAdic
    If .RecordCount > 0 Then
       .MoveFirst
       Do While Not .EOF
                MqueryP = MqueryP & "i" & !codigo & "=" & CCur(!Monto) & ""
            .MoveNext
          If Not .EOF Then MqueryP = MqueryP & ","
       Loop
    End If
End With

'Descuentos Adicionales

Dim MqueryD As String
MqueryD = ""
With pRsDesAdic
    If .RecordCount > 0 Then
       .MoveFirst
       Do While Not .EOF
                If !codigo = "13" And m_PagoQuincena = False Then
                  '---MqueryD = MqueryD & "d" & !codigo & "=" & "Round(" & CCur(!Monto) & "/" & intFactorCarga & ", 2)" & ""
                  MqueryD = MqueryD & "d" & !codigo & "=" & Round(CCur(!Monto) / 2, 2) & ""
                Else
                  MqueryD = MqueryD & "d" & !codigo & "=" & CCur(!Monto) & " "
                End If
                
                If !codigo = "13" And !Monto <> 0 Then s5taCargada = True
                
          .MoveNext
          If Not .EOF Then MqueryD = MqueryD & ","
       Loop
    End If
End With

Dim itemcosto As Integer
itemcosto = 1
Dim mcad As String
mcad = ""
With pRsCosto
If .RecordCount > 0 Then .MoveFirst
Do While Not .EOF
        If Not IsNull(!Monto) Then
           If !Monto <> 0 Then
                mcad = mcad & "ccosto" & Format(itemcosto, "0") & " = '" & !codigo & "'," & "porc" & Format(itemcosto, "0") & " = '" & Trim(Str(!Monto)) & "',"
           End If
        End If
    .MoveNext
   itemcosto = itemcosto + 1
Loop
End With
If Trim(mcad) <> "" Then
    mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
    mcad = mcad + ","
End If

Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql = Sql & " Update platemphist set " & mcad & "" & IIf(Trim(MqueryH) = "", "", MqueryH & ",") & MqueryP & "," & MqueryD
Sql = Sql & " where cia='" & m_CodCia & "' and placod='"
Sql = Sql & Trim(m_CodTrabajor) & "' and codauxinterno='"
Sql = Sql & Trim(m_CodAuxiInterno) & "' and proceso='"
Sql = Sql & Trim(m_TipoBoleta) & "' and fechaproceso='"
Sql = Sql & Format(m_FechaProceso, FormatFecha)
Sql = Sql & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
cn.Execute Sql, 64

'
Dim I As Integer
Dim MqueryI As String
MqueryI = ""
For I = 1 To 50
       'If I = 38 And (m_CodTrabajor = "EO212" Or m_CodTrabajor = "EO191") Then Stop
       Select Case m_TipoBoleta
               Case Is = "01" 'Normal
               
               If I = 2 And CargaPlanilla = True And B_ASIGFAM = True Then
               
                 Sql = F01(Trim(m_CodCia), Trim(m_CodTrabajor), Format(I, "00"), m_Status, m_año_proceso, m_mes_proceso, m_TipoTrabajor, "", m_CargoTrab, m_TopeHrsExtras, m_FechaProceso, "", CInt(m_NroSemana))
                 
               ElseIf I = 2 And CargaPlanilla = True And B_ASIGFAM = False Then
                  '  I = I + 1
                   ' Sql = F01(Trim(m_CodCia), Trim(m_CodTrabajor), Format(I, "00"), m_Status, m_año_proceso, m_mes_proceso, m_TipoTrabajor, "", m_CargoTrab, m_TopeHrsExtras, m_FechaProceso, "", CInt(m_NroSemana))
                   
                    Sql = "Select 0 as basico"
                    
                  Else
                  
                    Sql = F01(Trim(m_CodCia), Trim(m_CodTrabajor), Format(I, "00"), m_Status, m_año_proceso, m_mes_proceso, m_TipoTrabajor, "", m_CargoTrab, m_TopeHrsExtras, m_FechaProceso, "", CInt(m_NroSemana))
                    
               End If
                   
              ' End If
               Case Is = "02" 'Vacaciones
                    Sql = V01(Trim(m_CodCia), Trim(m_CodTrabajor), m_TopeHrsExtras, m_CargoTrab, m_año_proceso, m_mes_proceso, m_Status, m_TipoTrabajor, Format(I, "00"))
               Case Is = "03" 'Gratificaciones
                    Sql = G01(Trim(m_CodCia), Trim(m_CodTrabajor), m_TopeHrsExtras, m_CargoTrab, m_TipoTrabajor, m_año_proceso, m_mes_proceso, m_Status, Format(I, "00"), m_FechaProceso)
            Case Is = "04"
                Sql = F01(Trim(m_CodCia), Trim(m_CodTrabajor), Format(I, "00"), m_Status, m_año_proceso, m_mes_proceso, m_TipoTrabajor, "", m_CargoTrab, m_TopeHrsExtras, m_FechaProceso, Val(m_NroSemana))
            Case Is = "05"
                Sql = F01(Trim(m_CodCia), Trim(m_CodTrabajor), Format(I, "00"), m_Status, m_año_proceso, m_mes_proceso, m_TipoTrabajor, "", m_CargoTrab, m_TopeHrsExtras, m_FechaProceso, Val(m_NroSemana))
        End Select

    If Trim(Sql) <> "" Then
       cn.CursorLocation = adUseClient
       
       Dim Ri As ADODB.Recordset
    '   Sql = "SELECT H11,PLACOD FROM platemphist WHERE PLACOD ='" & m_CodTrabajor & "' AND YEAR(FECHAPROCESO)=2008 AND MONTH(FECHAPROCESO)=6"
       If fAbrRst(Ri, Sql) Then
          Ri.MoveFirst
          If IsNull(Ri(0)) Or Ri(0) = 0 Then
          Else
            If m_PagoQuincena = False And (I = 10 Or I = 11 Or I = 13 Or I = 21 Or I = 22 Or I = 23 Or I = 24) Then
                
               '---MqueryI = MqueryI & "i" & Format(I, "00") & " = " & "Round(" & Ri(0) & "/" & intFactorCarga & ", 2) ,"
                
               MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Round(Ri(0) / 2, 2) & ","
                
            Else
             MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Ri(0) & ","
            End If

          End If
       End If
       If Ri.State = 1 Then Ri.Close
    End If
Next

If MqueryI <> "" Then
   MqueryI = Mid(MqueryI, 1, Len(Trim(MqueryI)) - 1)
   Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql = Sql & " Update platemphist set " & MqueryI
   Sql = Sql & " where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
   Sql = Sql & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
   cn.Execute Sql
End If


'/*** CALCULA DEDUCCIONES  ***/


Dim MqueryCalD As String
MqueryCalD = ""
For I = 1 To 20
    'If I = 4 And m_CodTrabajor = "EO200" Then Stop
    'If I = 13 And m_CodTrabajor = "EE012" Then Stop
    If s5taCargada And I = 13 Then
        Sql$ = ""
    Else
        Sql$ = Me.F02(Format(I, "00"), m_CancelaGrabacion, m_CodAfp, m_Fecha_Jubilacion, m_Sn_Essaludvida, m_Sn_sindicato, m_CodCia, m_CodTrabajor, m_año_proceso, m_mes_proceso, m_TipoBoleta, m_Status, m_Manos, m_Sn_Quinta, m_TipoTrabajor, m_TopeHrsExtras, m_CargoTrab, m_Horas, m_NroSemana, pNumQuincena)
    End If
    
    If m_CancelaGrabacion = True Then
       m_VokDevengue = False
       MsgBox "Se Cancelo la Grabacion", vbCritical, "Calculo de Boleta"
       Sql$ = wCancelTrans
       cn.Execute Sql$
       Screen.MousePointer = vbDefault
       Exit Function
    End If
    If Sql$ <> "" Then
       If (fAbrRst(rs, Sql$)) Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
          Else
             If I = 11 Then
                For J = 1 To 5
                   MqueryCalD = MqueryCalD & "d" & Format(I, "00") & Format(J, "0") & " = " & rs(J - 1) & ","
                Next J
             ElseIf I = 13 And wtipodoc = False Then
                If rs(0) < 0 Then
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                Else
                   MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) / 2 & ","
                End If
             ElseIf I = 6 Then
                '2011.03.09 PARA QUE NO CALCULE NUEVAMENTE EL ESSALUD VIDA SI YA LO CALCULO EN ALGUN TIPO DE BOLETA
                Dim Cadena As String
                Dim rsVal As ADODB.Recordset
                Cadena = "select count(*) as Contador from plahistorico where cia = '" & wcia & "' and placod = '" & m_CodTrabajor & "' and status <> '*' and year(fechaproceso) = " & m_año_proceso & " and month(fechaproceso) = " & m_mes_proceso & " and d06<>0"
                Set rsVal = OpenRecordset(Cadena, cn)
                If Not rsVal.EOF Then
                    If Not rsVal.EOF Then
                        If rsVal!contador > 0 Then
                            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                        Else
                            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
                        End If
                    End If
                Else
                    'EN TEORIA NO DEBERIA SALTAR A ESTE LADO PERO EN CASO DE HACERLO SE VERA REFLEJADO EN LA BOLETA
                    'DE SER ASI MODIFICAR EL QUERY DE LA VARIABLE 'CADENA'
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
                End If
                
             Else
                If rs(0) < 0 Then
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                Else
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
                End If
             End If
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next I

If MqueryCalD <> "" Then
   MqueryCalD = Mid(MqueryCalD, 1, Len(Trim(MqueryCalD)) - 1)
   Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set " & MqueryCalD
   Sql$ = Sql$ & " where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
   cn.Execute Sql
   
   Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set d11=d111+d112+d113+d114+d115"
   Sql$ = Sql$ & " where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
   cn.Execute Sql
   
End If
'/*** CALCULO APORTACIONES ***/
Dim MqueryCalA As String
MqueryCalA = ""
For I = 1 To 20
    'If i = 1 Then Stop
    Sql$ = F03(Format(I, "00"), m_CodCia, m_CodArea, m_TipoBoleta, m_CodTrabajor, m_año_proceso, m_mes_proceso, m_Status, pm_TipoTrabajor)
    If Sql$ <> "" Then
       cn.CursorLocation = adUseClient
       Set rs = New ADODB.Recordset
       Set rs = cn.Execute(Sql$, 64)
       If rs.RecordCount > 0 Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
          Else
             MqueryCalA = MqueryCalA & "a" & Format(I, "00") & " = " & rs(0) & ","
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next
If MqueryCalA <> "" Then
   MqueryCalA = Mid(MqueryCalA, 1, Len(Trim(MqueryCalA)) - 1)
   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set " & MqueryCalA
   Sql$ = Sql$ & " where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & Trim(m_TipoBoleta) & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
   cn.Execute Sql
End If


Dim mi As String, md As String, ma As String

'Se modifico para que no calcule el campo h14 orden jcbr 14/05/2008
'Se modifico para que calcule    el campo h14 dias trabjados 25/06/2008 RHL

If pm_CalcularH14 Then
    mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
    md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
    ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
    Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    Sql$ = Sql$ & "update platemphist set h14=round(h01+h02+h03/8,0),totaling=" & mi & "," _
         & "totalded=" & md & "," _
         & "totalapo=" & ma & "," _
         & "totneto=(" & mi & ")-" & "(" & md & ")"
       Sql$ = Sql$ & " where cia='" & m_CodCia & "' and " & _
       " placod='" & Trim(m_CodTrabajor) & "' and " & _
       "codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
       Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
    cn.Execute Sql
Else

    mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
    md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
    ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
    Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    Sql$ = Sql$ & "update platemphist set h14=round((h01)/8,0),totaling=" & mi & "," _
         & "totalded=" & md & "," _
         & "totalapo=" & ma & "," _
         & "totneto=(" & mi & ")-" & "(" & md & ")"
       Sql$ = Sql$ & " where cia='" & m_CodCia & "' and " & _
       " placod='" & Trim(m_CodTrabajor) & "' and " & _
       "codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
       Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
       cn.Execute Sql
End If


'/*** jcms 27/05/08 calcular Deduccion Judicial si tiene PORCENTAJE***/
   If m_NuevaBoleta = True Then
       Dim Sqlr As String
       Sqlr = "select concepto,importe, status  from pladeducper "
       Sqlr = Sqlr & "where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & Trim(m_CodAuxiInterno) & "' and status='P'"
       Dim Rq As ADODB.Recordset
       If fAbrRst(Rq, Sqlr) Then
        
            Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
            Sql$ = Sql$ & " update platemphist set isnull(d05,0)=(isnull(totneto,0)+isnull(d09,0)+isnull(d07,0)+isnull(d12,0))*" & IIf(IsNull(Rq!importe), 0, Rq!importe)
            Sql$ = Sql$ & " where cia='" & m_CodCia & "' and "
            Sql$ = Sql$ & " placod='" & Trim(m_CodTrabajor) & "' and "
            Sql$ = Sql$ & " codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
            Sql$ = Sql$ & " and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
            cn.Execute Sql
            
            Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
            Sql$ = Sql$ & "update platemphist set totaling=" & mi & ","
            Sql$ = Sql$ & "totalded=" & md & ","
            Sql$ = Sql$ & "totalapo=" & ma & ","
            Sql$ = Sql$ & "totneto=(" & mi & ")-" & "(" & md & ")"
            Sql$ = Sql$ & " where cia='" & m_CodCia & "' and " & _
            Sql$ = Sql$ & " placod='" & Trim(m_CodTrabajor) & "' and "
            Sql$ = Sql$ & " codauxinterno='" & Trim(m_CodAuxiInterno) & "' and proceso='" & m_TipoBoleta & "' "
            Sql$ = Sql$ & " and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
            cn.Execute Sql

        
       End If
       Rq.Close
       Set Rq = Nothing
   End If
'/***


If m_PagoQuincena = True Or pNumQuincena = "2" Then
   Sql$ = "insert into plahistorico select * from platemphist"
Else
   Sql$ = "insert into plaquincena select * from platemphist"
End If
Sql$ = Sql$ & " where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & m_CodAuxiInterno & "' and proceso='" & m_TipoBoleta & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
cn.Execute Sql



   Dim ret As Long
   'Le pasamos a GetSafeArrayPointer el Arreglo y el valor de retorno
   GetSafeArrayPointer ArrDsctoCTACTE(), ret
   
   If ret > 0 Then
      'MsgBox " El array está dimensionado "
        If m_NuevaBoleta Then
            Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
            Sql$ = Sql$ & "select cia,placod,proceso,fechaproceso,semana,tipotrab,d07 from platemphist"
            Sql$ = Sql$ & " where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "' and codauxinterno='" & m_CodAuxiInterno & "' and proceso='" & m_TipoBoleta & "' "
            Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, Coneccion.FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
            
            If (fAbrRst(rs, Sql$)) Then
               If Not Descarga_ctaCte(m_CodCia, rs!PlaCod, m_CodAuxiInterno, rs!Proceso, Format(rs!FechaProceso, Coneccion.FormatFecha), rs!semana, rs!tipotrab, rs!d07) Then GoTo ErrCalc:
            End If
        End If
   Else
      'MsgBox " El array no está dimensionado "
   End If



Sql$ = "delete from platemphist where cia='" & m_CodCia & "' and placod='" & Trim(m_CodTrabajor) & "'"
Sql$ = Sql$ & " and cia='" & m_CodCia & "' and placod='" & m_CodTrabajor & "' and codauxinterno='" & m_CodAuxiInterno & "' and proceso='" & m_TipoBoleta & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(m_FechaProceso, FormatFecha) & "' and semana='" & m_NroSemana & "' and status='" & m_Status & "'"
cn.Execute Sql



cn.Execute "commit transaction", 64
Screen.MousePointer = 0

'Set pRsBoleta = m_RsBoleta.Clone
CalcularBoleta = True
Exit Function
ErrCalc:

cn.Execute "rollback transaction", 64
Screen.MousePointer = 0
CalcularBoleta = False
MsgBox "Error al generar la Boleta" & Chr(13) & ERR.Number & "-" & ERR.Description, vbCritical, "CalcularBoleta"
End Function
Private Sub Class_Initialize()
    Erase ArrDsctoCTACTE
End Sub
Public Function F01(ByVal pCodCia As String, ByVal pCodTrab As String, pCodConcepto As String, _
    ByVal pStatus As String, ByVal Paño As Integer, ByVal pMes As Integer, ByVal pTipTrab As String, ByVal pVacacion As String, _
    ByVal pCargo As String, ByVal pTopeHrsExt As String, ByVal pFechaProceso As String, ByVal pAltitud As String, Optional ByVal pSemana As Integer = 0) As String   'INGRESOS
    
Dim rsF01 As ADODB.Recordset
Dim mFactor As Currency
Dim nHijos As Integer
Dim RX As New ADODB.Recordset
Dim sSQL As String
Dim pSemanaPago As Integer
mFactor = 0
nHijos = 0
F01 = ""
Select Case pCodConcepto
       Case Is = "01" 'BASICO
            F01 = "select round((b.importe/factor_horas)*a.h01,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(pCodCia) & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & Trim(pCodConcepto) & "' and b.status<>'*'"
       Case Is = "02" 'ASIGNACION FAMILIAR
       
            sSQL = "SELECT semana_pago FROM PLACONSTANTE WHERE STATUS!='*' AND CIA='" & pCodCia & "' and codinterno='02' and tipomovimiento='02'"
            
            If (fAbrRst(RX, sSQL)) Then pSemanaPago = RX(0)
            
            If RX.State = 1 Then RX.Close

            If pTipTrab = "02" Then
                '*******codigo agregado giovanni 15092007***************************************
                'Call Recupera_Tipo_de_Tiempo(pCodCia, Trim(pCodTrab))

                'Select Case Reportes_Centrales.rs_RptCentrales_pub!descrip
                 '   Case "SEMANAL"
                  '      F01 = "select b.importe/(select count(semana) from plasemanas where cia='03'and year(fechaf)=" & paño & " and month(fechaf)=" & pmes & ") "
                   '     F01 = F01 & "as basico from platemphist a,plaremunbase b "
                    '    F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                     '   F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
                      '  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
                    'Case Else
                     '   If Semana_Calcular(pSemana, pSemanaPago, Year(FrmCabezaBol.Cmbfecha), pCodCia) Then
                      '      F01 = "select round((b.importe/factor_horas)*(factor_horas),2) as basico from platemphist a,plaremunbase b "
                       '     F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                        '    F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
                         '   F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
                        'End If
                'End Select
                'Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
                '*******************************************************************************
                'pSemanaPago = 4
                
                If Not pSemanaPago = 0 And XB_ASIGFAM Then
                
                    F01 = "select round((b.importe),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
                              
                End If
                
                If Semana_Calcular(pSemana, pSemanaPago, Trim(Str(Paño)), pCodCia) Then
                
                
                'F01 = "Select 0 as basico"
                '/factor_horas)*(factor_horas)
                
                    F01 = "select round((b.importe),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
                    
                Else
                
                ' FEC. MODIFICACION 20/06/2008
                ' SI EN EL CASO NO SE HA APLICADO LA ASIG. FAM. CARGARLA EN LA BOLETA
                ' MOD. POR RHL
                
                    Sql$ = " SELECT SUM(I02)AS_FAM " & _
                           "From PLAHISTORICO " & _
                           "WHERE CIA ='" & pCodCia & "'  " & _
                             "AND PLACOD ='" & pCodTrab & "'  " & _
                             "AND YEAR(FECHAPROCESO)=" & Paño & " " & _
                             "AND MONTH(FECHAPROCESO)=" & pMes & " " & _
                             "AND STATUS<>'*'"
                             
                     Dim rs_Semana_Calculo As ADODB.Recordset
                     
                     Set rs_Semana_Calculo = New ADODB.Recordset
            
                     rs_Semana_Calculo.Open Sql, cn, adOpenKeyset, adLockOptimistic
            
                   ' If rs_Semana_Calculo!AS_FAM = 0 Then
                    '
                     '   F01 = "select round((b.importe/factor_horas)*(factor_horas),2) as basico from platemphist a,plaremunbase b "
                      '  F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                       ' F01 = F01 & "and year(fechaproceso)=" & pAño & " and month(fechaproceso)=" & pMes & " "
                        'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
                    'End If
                        
                End If
            Else
            
            If m_PagoQuincena = False Then
            
                F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+a.h08+a.h02),2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
                
            Else
                F01 = "select round((b.importe),2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
                F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
            End If
            '/factor_horas)*(a.h01+a.h03+a.h08+a.h02)
            
                
            End If
            
       Case Is = "03" 'ASIGNACION MOVILIDAD
            ' CAMBIO DE H14 POR H01  {>MA<} 10/06/2007
            F01 = "select round((b.importe/factor_horas)*a.H01+A.H03,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
        
       Case Is = "04" 'BONIFICACION T. SERVICIO
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h02+a.h03+a.h04+a.h05+a.h12),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            'F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           
       Case Is = "05" 'INCREMENTO AFP 10.23%

            '*******************codigo Original**********************************
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & pCodTrab & "' "
            'F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"

            '************codigo agregado giovanni 05082007****************
            Select Case (pTipTrab)
                Case "01": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                Case "02": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            End Select
            '*************************************************************
            F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
            '*****************************************************************************************
            
       Case Is = "06" 'INCREMENTO AFP 3%
            '*****************codigo modificado giovanni 29082007**************************************
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & Trim(pCodCia) & "' and a.status='" & pstatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            'F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"

            '************codigo agregado giovanni 05082007****************
            Select Case Trim(pTipTrab)
                Case "01": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                Case "02": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            End Select
            '*************************************************************
            
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(pCodCia) & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
            '******************************************************************************************
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
            
       Case Is = "08" 'SOBRETASA (CONSTRUCCION CIVIL)
            Sql$ = "select factor from platasaanexo where cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
          
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select ROUND(" & mFactor & " *(a.h13),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               'F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h13),2) as basico from platemphist a,plaremunbase b "
               'F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & Trim(pCodTrab) & "' "
               'F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
               'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "09" 'DOMINICAL
            If pTipTrab <> "01" Then
               F01 = "select round(((b.importe/factor_horas)*h02),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               Else
               F01 = "select round(((b.importe/factor_horas)*h02),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "10" 'EXTRAS L-S
            
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
''               F01 = "select round((b.importe/factor_horas)* ((" & mFactor & " *(a.h10))+ a.h10),2) as basico from platemphist a,plaremunbase b "
''               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & pCodTrab & "' "
''               F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
''               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               '=====================================
               
                 'bts = convierte_cant
               
'                 F01 = "select round((((b.importe+" & bts & "+A.I07+A.I06)/factor_horas)* " & mFactor & ") *(a.h10),2) as basico,B.IMPORTE,FACTOR_HORAS from platemphist a,plaremunbase b "
'                 F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & Trim(pCodTrab) & "' "
'                 F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
'                 F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.status<>'*'"
                'If Trim(pCodTrab) = "RO007" Then Stop
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h10,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                 F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
                 
            End If
       Case Is = "11" 'EXTRAS D-F
       
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and  modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                'bts = convierte_cant
               
'               F01 = "select round((((b.importe+" & bts & ")/factor_horas)* " & mFactor & ") *(a.h11),2) as basico from platemphist a,plaremunbase b "
'               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pstatus & "' and a.placod='" & pCodTrab & "' "
'               F01 = F01 & "and year(fechaproceso)=" & paño & " and month(fechaproceso)=" & pmes & " "
'               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h11,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND B.CIA ='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"

            End If
       Case Is = "12" 'FERIADOS
            F01 = "select round((b.importe/factor_horas)*a.h03,2) as basico,A.H03 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "13" 'REINTEGROS
       Case Is = "14" 'VACACIONES (CONSTRUCCION CIVIL)
            If pVacacion = "S" Then
               Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and  modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "15" 'GRATIFICACION
       Case Is = "29" 'HORARIO DE VERANO
       Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*' AND CIA ='" & pCodCia & "'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                
                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h15,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' AND B.CIA='" & pCodCia & "'"
       End If
       Case Is = "17" 'ASIGNACION ESCOLAR
            If pTipTrab = "01" Then
            End If
            If pTipTrab = "02" Then
            End If
            If pTipTrab = "05" Then
               Sql$ = "select ultima from plasemanas where cia='" & pCodCia & "' and semana='" & Format(pSemana, "00") & "' and ano='" & Paño & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  If rsF01!ultima = "S" Then
                     Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
                     If (fAbrRst(rsF01, Sql$)) Then
                        nHijos = Numero_Hijos(Trim(pCodTrab), "S", "S", pFechaProceso, 18)
                        mFactor = rsF01!factor
                        If rsF01.State = 1 Then rsF01.Close
                        F01 = "select round((((b.importe/factor_horas)*8)* " & mFactor & ")/12 * " & nHijos & ",2) as basico from platemphist a,plaremunbase b "
                        F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                        F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                        F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
                     End If
                  End If
               End If
            End If
       Case Is = "18" 'UTILIDADES
       Case Is = "20" 'BUC
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
            
        Case Is = "21" '3RO HORAS EXTRAS
                Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
                mFactor = rsF01!factor
                If rsF01.State = 1 Then rsF01.Close

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h17,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
                
            End If

       Case Is = "22" 'BONIF. CONTACTO AGUA
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h19),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
            
       Case Is = "23" 'BONIF.POR ALTITUD
            If pAltitud = "S" Then
               Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round((" & mFactor & " / 8) * (a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "24" 'BONIF. TURNO NOCHE
       
         Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
         
            If (fAbrRst(rsF01, Sql$)) Then
                mFactor = rsF01!factor
                If rsF01.State = 1 Then rsF01.Close

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h18,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
                
            End If
            
            '
            
            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
            
            'If (fAbrRst(rsF01, Sql$)) Then
             '  mFactor = rsF01!factor
              ' If rsF01.State = 1 Then rsF01.Close
               
               'F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.H19),2) as basico from platemphist a,plaremunbase b "
               'F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               'F01 = F01 & "and year(fechaproceso)=" & pAño & " and month(fechaproceso)=" & pMes & " "
               'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               
               'F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h19,2)) as basico"
               'F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & pAño & " and MONTH(fechaproceso)=" & pMes & ") "
               'F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
               
            'End If
        
       Case Is = "38" '//** jcms BONIF.PROD. 12/06/2008
       
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
'                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h10,2)) as basico"
'                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & pAño & " and MONTH(fechaproceso)=" & pMes & ") "
'                 F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02')"
                 
                 

                'EXTRAS L-S
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h24,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                 F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
                 
                 
            End If
            
       Case Is = "25" 'H.E. NOCHE
       
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            
            If (fAbrRst(rsF01, Sql$)) Then
                mFactor = rsF01!factor
                If rsF01.State = 1 Then rsF01.Close

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h19,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Paño & " and MONTH(fechaproceso)=" & pMes & ") "
                F01 = F01 & " WHERE A.CIA='" & pCodCia & "' AND A.PLACOD='" & Trim(pCodTrab) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
                '
            End If
            
       Case Is = "26" 'H.E. HASTA ONCEAVA HORA
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h22),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "27" 'EXTRAS 3PRA L-S
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
            
               mFactor = rsF01!factor
               
               If rsF01.State = 1 Then rsF01.Close
               
                F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h23),2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               
            End If
       Case Is = "28" 'EXT. NOCHE 2PR L-S
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h24),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "35" 'EXT. NOCHE 3PRA L-S (29)
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "'  and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h25),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "30" 'SOBRETASA NOCHE(CONSTRUCCION CIVIL)
            Sql$ = "select factor from platasaanexo where  cia='" & pCodCia & "' and modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipTrab & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h07),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
      Case Is = "36"
      
            F01 = "select round((b.importe/factor_horas)*a.h23,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(pCodCia) & "' and a.status='" & pStatus & "' and a.placod='" & Trim(pCodTrab) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
        
End Select

End Function


Public Function V01(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal pTopeHrsExt As String, ByVal pCargo As String, ByVal Paño As Integer, ByVal pMes As Integer, _
                    ByVal pStatus As String, ByVal pTipoTrab As String, ByVal pCodConcepto As String) As String
'INGRESOS
Dim rsV01 As ADODB.Recordset
Dim mFactor As Currency
Dim fACTORfAMILIA As String

'FECHA DE MODIFICACION 08/01/2008
'If pcodcia = "05" Then fACTORfAMILIA = 2 Else fACTORfAMILIA = 1

If pCodCia = "05" Then fACTORfAMILIA = 1 Else fACTORfAMILIA = 1

mFactor = 0
V01 = ""
Select Case pCodConcepto
       Case Is = "02" 'ASIGNACION FAMILIAR
            V01 = "select round((b.importe/factor_horas)*(a.h12*" & fACTORfAMILIA & "),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
       Case Is = "04" 'BONIFICACION T. SERVICIO
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
       Case Is = "05" 'INCREMENTO AFP 10.23%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
       Case Is = "06" 'INCREMENTO AFP 3%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
       Case Is = "14" 'VACACIONES
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
            V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "20" 'BUC
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipoTrab & "' and cargo='" & Trim(pCargo) & "' and status<>'*'"
            If (fAbrRst(rsV01, Sql$)) Then
               mFactor = rsV01!factor
               If rsV01.State = 1 Then rsV01.Close
               V01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h12),2) as basico from platemphist a,plaremunbase b "
               V01 = V01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
               V01 = V01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
               V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "29"
       
        'MODIFICADO POR RICARDO HINOSTROZA
        'AGREGAR CALCULO DE HORARIO VACACIONAL
        'FEC MODIFICACION 10/04/2008
        
             V01 = "SELECT ROUND(SUM(I29)/6,2)  AS BASICO FROM PLAHISTORICO WHERE CIA ='" & pCodCia & "' AND STATUS<>'*' AND PLACOD='" & pCodTrab & "' AND YEAR(FECHAPROCESO) =" & Paño & " AND (MONTH(FECHAPROCESO) BETWEEN 1 AND 3)"
       
End Select
'Debug.Print pcodconcepto
'Debug.Print SQL$
End Function


Public Function G01(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal pTopeHrsExt As String, ByVal pCargo As String, _
                    ByVal pTipoTrab As String, ByVal Paño As Integer, ByVal pMes As Integer, ByVal pStatus As String, pCodConcepto As String, ByVal pFechaProceso As String) As String      'INGRESOS
Dim rsG01 As ADODB.Recordset
Dim mFactor As Currency
Dim mh As Integer
mFactor = 0
If Val(Mid(pFechaProceso, 4, 2)) = 7 Then mh = 7 Else mh = 5
G01 = ""
If pTipoTrab <> "05" Then
    Select Case pCodConcepto
           Case Is = "02" 'ASIGNACION FAMILIAR
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           Case Is = "04" 'BONIFICACION T. SERVICIO
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           Case Is = "05" 'INCREMENTO AFP 10.23%
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           Case Is = "06" 'INCREMENTO AFP 3%
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           Case Is = "07" 'BONIFICACION COSTO DE VIDA
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & pCodConcepto & "' and b.status<>'*'"
           Case Is = "15" 'GRATIFICACION
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
    End Select
Else
   Select Case pCodConcepto
          Case Is = "15" 'GRATIFICACION
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(pTopeHrsExt) & "' and codinterno='" & pCodConcepto & "' and tipotrab='" & pTipoTrab & "' and cargo='" & Trim(pCargo) & "' and status<>'*'"
               If (fAbrRst(rsG01, Sql$)) Then
                  mFactor = rsG01!factor
                  If rsG01.State = 1 Then rsG01.Close
                  G01 = "select round(((((b.importe/factor_horas)*8)* " & mFactor & ") / " & mh & ") * a.h14 ,2) as basico from platemphist a,plaremunbase b "
                  G01 = G01 & "where a.cia='" & pCodCia & "' and a.status='" & pStatus & "' and a.placod='" & pCodTrab & "' "
                  G01 = G01 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " "
                  G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
   End Select
End If
End Function

'DEDUCCIONES

Public Function F02(ByVal concepto As String, ByRef mcancel As Boolean, ByVal pCodAfp As String, ByVal pFechaJubilacion As String _
, ByVal pSn_essaludvida As Byte, ByVal pSn_sindicato As Byte _
, ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As Integer, ByVal pMes As Integer _
, ByVal pTipoBol As String, ByVal pStatus As String, ByVal pManos As Integer, ByVal pSn_Quinta As Boolean _
, ByVal pTipoTrab As String, ByVal pTopeHrsExt As String, ByVal pCargoTrab As String _
, ByVal pHoras As Integer, ByVal pNroSemana As Integer, pNumQuincena As String) As String


Dim rsF02 As ADODB.Recordset
Dim rsImporte As ADODB.Recordset
Dim rsF02afp As ADODB.Recordset
Dim F02str As String
Dim rsTope As ADODB.Recordset
Dim mFactor As Currency
Dim mperiodoafp As String
Dim vNombField As String
Dim mtope As Currency
Dim mincremento As Currency
Dim difmincremento As Currency
Dim mproy As Currency
Dim MUIT As Currency
Dim mgra As Integer
Dim msemano As Integer
Dim mpertope As Integer
Dim J As Integer
Dim conceptosremu As String
Dim snmanual As Byte, snfijo As Byte
Dim cptoincrementos As String
Dim IMPORTEESVMES As Currency
Dim Porc_dEsSalud As Double
Dim Porc_EPS As Double
Dim Afecto_EPS As Boolean

Dim macui As Currency
Dim macus As Currency
Dim importe As Currency

mFactor = 0
F02 = ""
mtope = 0

Porc_dEsSalud = 0
Porc_EPS = 0
Afecto_EPS = False

'---
Dim intFactorCarga As Integer
'intFactorCarga = IIf(wtipodoc = True, 1, IIf(pNumQuincena = "1", 2, 1))
        

If (concepto <> "04" Or Trim(pCodAfp) = "01" Or Trim(pCodAfp) = "" Or Trim(pCodAfp) = "02") And concepto <> "11" And concepto <> "13" Then  'SIN AFP
   If Not IsDate(pFechaJubilacion) Then
    Sql$ = "select deduccion,adicional,status from placonstante where cia='" & pCodCia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and deduccion<>0 and status<>'*'"
  
    If (fAbrRst(rsF02, Sql$)) Then
       If Not IsNull(rsF02!deduccion) Then
          If rsF02!deduccion <> 0 Then mFactor = rsF02!deduccion: snmanual = IIf(rsF02!adicional = "S", 0, 1): snfijo = IIf(rsF02!status = "F", 1, 0)
       End If
    End If
    
    If rsF02.State = 1 Then rsF02.Close
    
    Dim sw As Boolean
    sw = True
    
    If snfijo = 1 And snmanual = 0 And ((pSn_essaludvida = 1 And concepto = "06") Or (pSn_sindicato = 1 And concepto = "08")) Then
        
        'Call Acumula_Mes(concepto, "D")
        
        Call Acumula_Mes(pCodCia, pCodTrab, Paño, pMes, concepto, "D", macui, macus, pTipoBol)
         
        If wtipodoc = True Then
            If macui = mFactor Then
                F02 = " SELECT " & 0
            Else
                F02 = " SELECT " & mFactor
            End If
        Else
            If macui = mFactor Then
                F02 = " SELECT " & 0
            Else
               'Solo para 1 Quincena intFactorCarga = 2
               F02 = " SELECT " & mFactor / 2
              'F02 = " SELECT " & mFactor & "/" & intFactorCarga
            End If
        End If
        sw = False
    End If
    
    If mFactor <> 0 And sw Then
    
       Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='D' and tboleta='" & pTipoBol & "'  and  codigo='" & concepto & "' and status<>'*'"
       If (fAbrRst(rsF02, Sql$)) Then
          rsF02.MoveFirst
          F02str = ""
          Do While Not rsF02.EOF
             F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
             rsF02.MoveNext
          Loop
          F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
          If rsF02.State = 1 Then rsF02.Close
        
           Call Acumula_Mes(pCodCia, pCodTrab, Paño, pMes, concepto, "D", macui, macus, pTipoBol)
        
          F02 = "select round(((" & F02str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as deduccion from platemphist "
          F02 = F02 & "where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' "
          F02 = F02 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
       End If
    End If
   End If
ElseIf Trim(concepto) = "11" And Trim(pCodAfp) <> "" Then 'AFP
   If Not IsDate(pFechaJubilacion) Then
   If Trim(pCodAfp) = "01" Or Trim(pCodAfp) = "02" Then GoTo AFP
    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='D' and tboleta='" & pTipoBol & "'  and  codigo='" & concepto & "' and status<>'*'"

    If (fAbrRst(rsF02, Sql$)) Then
       rsF02.MoveFirst
       F02str = ""
       Do While Not rsF02.EOF
          F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
          
          rsF02.MoveNext
       Loop
       
       F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
       
       If rsF02.State = 1 Then rsF02.Close
       mperiodoafp = Format(Paño, "0000") & Format(pMes, "00")
       Sql$ = "select afp01,afp02,afp03,afp04,afp05,tope from  plaafp where periodo='" & mperiodoafp & "' and codafp='" & pCodAfp & "' and status<>'*' and cia='" & pCodCia & "'"
    
       If Not (fAbrRst(rsF02, Sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para AFP", vbCritical, "Calculo de Boleta"
          mcancel = True
          'Exit Function
       End If
       
       Sql$ = Acumula_Mes_Afp(concepto, "D", pCodCia, Paño, pMes, macui, macus, pCodTrab)
       
       If (fAbrRst(rsF02afp, Sql$)) Then
          For J = 1 To 5
              
              vNombField = " as D11" & Format(J, "0")
              mFactor = rsF02(J - 1)
              
              If J = 2 Then
                 If pManos > 64 Then mFactor = 0
                
                 Call Acumula_Mes_Afp112(pCodCia, concepto, "D", macui, macus, pTipoBol, pCodTrab, Paño, pMes)
                 
                 mtope = macui
                 Sql$ = "select " & F02str & " as tope from platemphist where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' " _
                      & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
                      
                 If (fAbrRst(rsTope, Sql$)) Then mtope = mtope + rsTope!tope
                 
                 If wtipodoc = False Then mtope = mtope + rsTope!tope
                 
                 '---If wtipodoc = False And pNumQuincena = "1" Then mtope = mtope + rsTope!tope'
                 
                 If mtope > rsF02!tope Then mtope = rsF02!tope
                                 
                 If rsTope.State = 1 Then rsTope.Close
                 
                 If wtipodoc = False Then
                    F02 = F02 & "round((((" & mtope & ") * " & mFactor & " /100)-" & macus & ")/2, 2) "
                    '---F02 = F02 & "round((((" & mtope & ") * " & mFactor & " /100)-" & macus & ")/" & intFactorCarga & ", 2) "
                 Else
                    F02 = F02 & "round(((" & mtope & ") * " & mFactor & " /100)-" & macus & ",2) "
                 End If
                 F02 = F02 & vNombField & ","
              Else
                 If Not IsNull(rsF02afp(0)) Then
                    If wtipodoc Then
                        F02 = F02 & "round(((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & ",2) "
                        F02 = F02 & vNombField & ","
                    Else
                        F02 = F02 & "round((((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & "),2) "
                        F02 = F02 & vNombField & ","
                    End If
                 Else
                    If wtipodoc Then
                        F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                        F02 = F02 & vNombField & ","
                    Else
                        F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                        F02 = F02 & vNombField & ","
                    End If
                 End If
              End If
          Next J
          If rsF02afp.State = 1 Then rsF02afp.Close
          If rsF02.State = 1 Then rsF02.Close
       End If
       
       F02 = Mid(F02, 1, Len(Trim(F02)) - 1)
       F02 = "select " & F02
       F02 = F02 & " from platemphist "
       F02 = F02 & "where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' "
       F02 = F02 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
AFP:
    End If
   End If
               'de 13 lo cambie a 131 para que no entre
ElseIf concepto = "13" And pTipoBol <> "03" Then  'Quinta Categoria

    'PREGUNTAMOS SI ESTA AFECTO O NO
    
    'If pCodTrab = "EE012" Then Stop
    
    If pTipoBol = "04" Then GoTo quinta
    If pTipoBol = "05" Then GoTo quinta
    If Not pSn_Quinta Then GoTo quinta
    
   Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='D' and tboleta='" & pTipoBol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF02, Sql$)) Then
      rsF02.MoveFirst
      F02str = ""
      conceptosremu = ""
      cptoincrementos = ""
      Do While Not rsF02.EOF
         F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
         rsF02.MoveNext
      Loop
      
      F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      MUIT = 0
      mtope = 0
      mpertope = 0
      mgra = 0
      msemano = 0
      mincremento = 0
      If pTipoTrab <> "01" Then
         Sql$ = "select isnull(max(semana),0) as ultima from plasemanas where cia='" & pCodCia & "' and ano='" & Format(Paño, "0000") & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mpertope = rsF02(0): msemano = rsF02(0)
      Else
         If pMes > 6 Then mpertope = 12 Else mpertope = 13
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      'ACUMULADO DE TODOS INGRESOS
      
      Call Acumula_Ano(concepto, "D", macui, macus, pCodCia, pCodTrab, Paño, pMes)
      
      mtope = macui
      
      'OBTENER EL INCREMENTO
      If Len(Trim(cptoincrementos)) > 0 Then
        Sql$ = "select " & cptoincrementos & " as tope from platemphist where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' " _
             & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
        If (fAbrRst(rsF02, Sql$)) Then mincremento = rsF02!tope
        If rsF02.State = 1 Then rsF02.Close
      End If
      
      Sql$ = "select " & F02str & " as tope from platemphist where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' " _
           & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
      
      If (fAbrRst(rsF02, Sql$)) Then mtope = mtope + rsF02!tope
      
      If rsF02.State = 1 Then rsF02.Close
     
        Sql$ = "select concepto,moneda,sum((importe/factor_horas)) as base, SUM(importe) as importe from plaremunbase a,plaafectos b where a.cia='" & pCodCia & "' " _
           & "and placod='" & pCodTrab & "' and a.status<>'*' and b.tipo='D' AND B.CIA='" & pCodCia & "' and b.codigo='" & concepto & "' and b.tboleta='" & pTipoBol & "' and b.status<>'*' and a.concepto=b.cod_remu " _
           & "Group By Placod,a.concepto,a.moneda"
      
      If (fAbrRst(rsF02, Sql$)) Then
      importe = 0
        Do While Not rsF02.EOF
            mproy = mproy + rsF02!base
            importe = importe + rsF02!importe
            rsF02.MoveNext
        Loop

      End If
      
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select uit from plauit where ano='" & Format(Paño, "0000") & "' and moneda='S/.' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then MUIT = rsF02!uit
      If rsF02.State = 1 Then rsF02.Close
      
      If pTipoTrab = "05" Then
         Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & pTopeHrsExt & "' and codinterno='20' and tipotrab='" & pTipoTrab & "' and cargo='" & pCargoTrab & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mFactor = rsF02!factor
         If rsF02.State = 1 Then rsF02.Close
      
         Sql$ = "select concepto,moneda,importe/factor_horas as base from plaremunbase where cia='" & pCodCia & "' and placod='" & pCodTrab & "' and concepto='01' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mproy = mproy + Round(rsF02!base * mFactor, 2)
         If rsF02.State = 1 Then rsF02.Close
      End If
          
          
    'PARA EL CALCULO DE EPS
    Dim rsTemporal As ADODB.Recordset
    Cadena = "SELECT APORTACION FROM PLACONSTANTE WHERE CIA = '" & pCodCia & "' AND TIPOMOVIMIENTO = '03' AND STATUS != '*' AND CODINTERNO = '01'"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    If Not rsTemporal.EOF Then
        Porc_dEsSalud = Val(rsTemporal!aportacion)
    Else
        Porc_dEsSalud = 0
    End If
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    
    Cadena = "SELECT ISNULL(APORTACION_EPS,0) AS APORTACION_EPS FROM CIA WHERE COD_CIA = '" & pCodCia & "' AND STATUS != '*'"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    If Not rsTemporal.EOF Then
        Afecto_EPS = CBool(rsTemporal!Aportacion_EPS)
    Else
        Afecto_EPS = False
    End If
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    If Afecto_EPS Then
        Cadena = "SELECT ISNULL(CODIGO_EPS,'') AS CODIGO_EPS FROM PLANILLAS WHERE CIA = '" & pCodCia & "' AND PLACOD = '" & pCodTrab & "' AND STATUS != '*' AND TIPOTRABAJADOR = '" & pTipoTrab & "'"
        Set rsTemporal = OpenRecordset(Cadena, cn)
        If Not rsTemporal.EOF Then
            If rsTemporal!codigo_eps <> "" Then
                Cadena = "SELECT ISNULL(IMPORTE,0) AS IMPORTE FROM MAESTROS_2 WHERE RIGHT(CIAMAESTRO,3) = '143' AND STATUS != '*' AND RTRIM(ISNULL(CODSUNAT,'')) != '' AND COD_MAESTRO2 = '" & Trim(rsTemporal!codigo_eps) & "'"
                If rsTemporal.State = adStateOpen Then rsTemporal.Close
                Set rsTemporal = OpenRecordset(Cadena, cn)
                If Not rsTemporal.EOF Then
                    Porc_EPS = Val(rsTemporal!importe)
                Else
                    Porc_EPS = 0
                End If
            Else
                Porc_EPS = 0
            End If
        Else
            Porc_EPS = 0
        End If
    End If
      
      If pTipoTrab = "01" Then
         
         If wtipodoc = True Then
            If Afecto_EPS Then
                If pMes < 12 Then mproy = ((mproy * pHoras) + mincremento) * (mpertope - pMes + 1) + IIf(pMes > 6, (importe * Porc_EPS) / 100, ((importe * Porc_EPS) / 100) * 2) Else mproy = 0
            Else
                If pMes < 12 Then mproy = ((mproy * pHoras) + mincremento) * (mpertope - pMes + 1) + IIf(pMes > 6, (importe * Porc_dEsSalud) / 100, ((importe * Porc_dEsSalud) / 100) * 2) Else mproy = 0
            End If
         Else
            If pMes = 12 Then
               mproy = (((mproy * pHoras) + difmincremento) * (mpertope - pMes)) + Round((((mproy * pHoras) + difmincremento) / 2), 2)
            Else
              If Afecto_EPS Then
                mproy = (((mproy * pHoras) + difmincremento) * (mpertope - pMes + 1)) + Round((((mproy * pHoras) + difmincremento) / 2), 2) + (IIf(pMes > 6, (importe * Porc_EPS) / 100, ((importe * Porc_EPS) / 100) * 2))
              Else
                mproy = (((mproy * pHoras) + difmincremento) * (mpertope - pMes + 1)) + Round((((mproy * pHoras) + difmincremento) / 2), 2) + (IIf(pMes > 6, (importe * Porc_dEsSalud) / 100, ((importe * Porc_dEsSalud) / 100) * 2))
              End If
            
            End If
         End If
      Else
                   
         mgra = Busca_Grati(pCodCia, pCodTrab, Paño, pMes)
         
         If pTipoTrab = "05" Then
            Sql$ = "select importe/factor_horas as base,b.factor  from plaremunbase a,platasaanexo b where a.cia='" & pCodCia & "' and a.placod='" & pCodTrab & "'  and a.concepto='01' " _
                 & "and a.status<>'*' and b.cia='" & pCodCia & "' and b.tipomovimiento='01' and b.codinterno='15' and b.status<>'*' and b.tipotrab='" & pTipoTrab & "' and b.cargo='" & pCargoTrab & "'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = ((mproy * pHoras) + mincremento) * (mpertope - Val(pNroSemana)) + ((rsF02!base * 8) * (rsF02!factor * mgra)) + ((rsF02!base * 8) * (mpertope - Val(pNroSemana)))
            If rsF02.State = 1 Then rsF02.Close
         Else
         
         'MODIFICADO POR Ricardo Hinostroza
         'Fecha:18/10/2008
         'Motivo:calculo de boleta
         
            If Not pTipoBol = "02" Then
                Sql$ = "select importe/factor_horas as base from plaremunbase where Cia='" & pCodCia & "' and placod='" & pCodTrab & "'  and concepto='01' and status<>'*'"
                If (fAbrRst(rsF02, Sql$)) Then
                    If Afecto_EPS Then
                        mproy = ((mproy * pHoras) + mincremento) * (mpertope - Val(pNroSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(pNroSemana)) + IIf(pMes > 6, (importe * Porc_EPS) / 100, ((importe * Porc_EPS) / 100) * 2)
                    Else
                        mproy = ((mproy * pHoras) + mincremento) * (mpertope - Val(pNroSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(pNroSemana)) + IIf(pMes > 6, (importe * Porc_dEsSalud) / 100, ((importe * Porc_dEsSalud) / 100) * 2)
                    End If
                End If
                If rsF02.State = 1 Then rsF02.Close
            End If

         End If
      End If
      
      mtope = mtope + mproy

      If mtope > Round(MUIT * 7, 2) Then
         mtope = mtope - Round(MUIT * 7, 2)
         Select Case mtope
                Case Is < (Round(MUIT * 27, 2) + 1)
                     mFactor = Round(mtope * 0.15, 2)
                Case Is < (Round(MUIT * 54, 2) + 1)
                     mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
                Case Else
                     mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
         End Select
         
         If pTipoTrab = "01" Then
            If wtipodoc = True Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & pMes & " + 1), 2)"
            Else
               F02 = "Select Round(((" & mFactor & " - " & macus & ") / (12 - " & pMes & " + 1)), 2)"
            End If
         Else
            If pTipoBol = "02" Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & pMes & " + 1), 2)"
            Else
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (" & msemano & " - " & Val(pNroSemana) & " + 1), 2)"
            End If
         End If
      End If
   End If
quinta:
   
ElseIf concepto = "15" Then
       MsgBox "666"
End If
macui = 0: macus = 0
End Function
Private Sub Acumula_Mes(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As Integer, ByVal pMes As Integer, pCodConcepto As String, tipo As String, _
 ByRef Pmacui As Currency, ByRef Pmacus As Currency, pTipoBol As String)

Dim mcad As String
Dim SqlAcu As String

Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset

Dim mtb As String
Dim tope As Integer
Pmacui = 0: Pmacus = 0
Dim I As Integer

If pTipoBol = "01" Or pTipoBol = "02" Or pTipoBol = "04" Then
    tope = 3
ElseIf pTipoBol = "03" Then
    tope = 1
End If
For I = 1 To tope
    If pTipoBol = "01" Then
        If I = 1 Then mtb = "01"
    ElseIf pTipoBol = "03" Then
        If I = 1 Then mtb = "03"
    ElseIf pTipoBol = "04" Then
        If I = 1 Then mtb = "01"
    End If
    
    'Modificacion 15/02/2010
     
     If pTipoBol = "02" And pCodConcepto = "01" And tipo = "A" And I = 1 Then
       mtb = "01"
     End If
    
    'Fin Modificacion 15/02/2010
    
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "04"
        
    If pCodConcepto = "06" And tipo = "D" Then
        Sql$ = "select 'D06' AS COD_REMU"
    Else
        Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    End If
    
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
        If pCodConcepto = "06" And tipo = "D" Then
            mcad = mcad & Trim(RsAcumula!cod_remu) & "+"
        Else
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
        End If
          RsAcumula.MoveNext
       Loop
       
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(pCodConcepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & pCodTrab & "' "
        
        SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " and proceso='" & mtb & "'"
       
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing
End Sub

Private Sub Acumula_Mes_ANTERIOR(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As Integer, ByVal pMes As Integer, pCodConcepto As String, tipo As String _
, ByRef Pmacui As Currency, ByRef Pmacus As Currency)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Pmacui = 0: Pmacus = 0
Dim I As Integer
For I = 1 To 3
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If pCodConcepto = "06" And tipo = "D" Then
        Sql$ = "select 'D06' AS COD_REMU"
    Else
        Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    End If
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
        If pCodConcepto = "06" And tipo = "D" Then
            mcad = mcad & Trim(RsAcumula!cod_remu) & "+"
        Else
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
        End If
          RsAcumula.MoveNext
       Loop
       
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(pCodConcepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & pCodTrab & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing
End Sub

Private Function Acumula_Mes_Afp(ByVal pCodConcepto As String, pCodTipo As String, ByVal pCodCia As String, ByVal Paño As String, ByVal pMes As String, ByRef Pmacui As Currency, ByRef Pmacus As Currency, ByVal pCodTrab As String) As String

Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset

Pmacui = 0: Pmacus = 0
    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & pCodTipo & "' and tboleta='01'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d111) as ded1, " _
              & "sum(d112) as ded2, " _
              & "sum(d113) as ded3, " _
              & "sum(d114) as ded4, " _
              & "sum(d115) as ded5 "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & pCodTrab & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','04','05')"
      
      'PROCESO = 03 NO CONSIDERADO HASTA 2014
      'SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
    
    End If
If RsAcumula.State = 1 Then RsAcumula.Close
Acumula_Mes_Afp = SqlAcu
End Function

'Fecha Modificación : 30/07/2008
'Agregado por Ricardo Hinostroza
'Acumular SNP para Ajustar Aporte del AFP

Private Function Acumula_Snp(ByVal pMes As String, ByVal pCodTrab As String, ByVal Paño As String, _
                            ByVal pCodCia As String) As String
Dim SqlAcu As String


SqlAcu = " Select  sum(d04)*-1 as d04 " & _
         " From plahistorico " & _
         " where cia='" & pCodCia & "' " & _
         " and status<>'*' " & _
         " and placod='" & pCodTrab & "' " & _
         " and year(fechaproceso)=" & Paño & " " & _
         " and month(fechaproceso)=" & pMes & " " & _
         " and proceso in('01','02','03','04','05')"

Acumula_Snp = SqlAcu

End Function

Private Sub Acumula_Mes_Afp112(ByVal pCodCia As String, ByVal pCodConcepto As String, pTipo As String, ByRef Pmacui As Currency, ByRef Pmacus As Currency, ByVal pTipoBol As String, ByVal pCodTrab As String, ByVal Paño As String, ByVal pMes As String)

Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim tope As Integer
Pmacui = 0: Pmacus = 0
Dim I As Integer
If pTipoBol = "01" Or pTipoBol = "02" Then
    tope = 2
ElseIf pTipoBol = "03" Then
    tope = 1
End If
For I = 1 To tope
    
    If pTipoBol = "01" Then
        If I = 1 Then mtb = "01"
        If I = 2 Then mtb = "02"
    ElseIf pTipoBol = "02" Then
        If I = 1 Then mtb = "00"
        If I = 2 Then mtb = "02"
    ElseIf pTipoBol = "03" Then
        If I = 1 Then mtb = "03"
    End If

    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & pTipo & "' and tboleta='" & mtb & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d112) as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & pCodTrab & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing
End Sub



Private Sub Acumula_Mes_Afp112_ANTERIOR(ByVal pCodCia As String, ByVal pCodConcepto As String, pTipo As String, ByRef Pmacui As Currency, ByRef Pmacus As Currency, ByVal pTipoBol As String, ByVal pCodTrab As String, ByVal Paño As String, ByVal pMes As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Pmacui = 0: Pmacus = 0
Dim I As Integer
For I = 1 To 3
    If pTipoBol = "02" And I <> 2 Then I = I + 1
    If pTipoBol <> "02" And I = 2 Then I = I + 1
    If I > 3 Then Exit For
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & pTipo & "' and tboleta='" & mtb & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d112) as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & Trim(pCodTrab) & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing
End Sub

Private Sub Acumula_Ano(ByVal pCodConcepto As String, ByVal pTipo As String, ByRef Pmacui As Currency, ByRef Pmacus As Currency, ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As String, ByVal pMes As String)

Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim I As Integer
Pmacui = 0: Pmacus = 0

For I = 1 To 5
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If I = 4 Then mtb = "04"
    If I = 5 Then mtb = "10"
    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & pTipo & "' and tboleta='" & mtb & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop

       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(pTipo) & Trim(pCodConcepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & pCodTrab & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)<=" & pMes & " and proceso='" & mtb & "'"

       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If
Next

If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing

End Sub


Private Sub Acumula_Ano_ANTERIOR(ByVal pCodConcepto As String, ByVal pTipo As String, ByRef Pmacui As Currency, ByRef Pmacus As Currency, ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As String, ByVal pMes As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Pmacui = 0: Pmacus = 0
Dim I As Integer
For I = 1 To 4
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If I = 4 Then mtb = "04"
    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & pTipo & "' and tboleta='" & mtb & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop

       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(pTipo) & Trim(pCodConcepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & Trim(pCodTrab) & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)<=" & pMes & " and proceso='" & mtb & "'"

       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing
End Sub

Private Function Busca_Grati(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As String, ByVal pMes As String) As Integer

Dim rsGrati As New Recordset
Select Case pMes
       Case Is = 1, 2, 3, 4, 5, 6
            Busca_Grati = 2
       Case Is = 7
            Sql$ = "select * from plahistorico where cia='" & pCodCia & "' and placod='" & pCodTrab & "' and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 1 Else Busca_Grati = 2
                 
       Case Is = 8, 9, 10, 11
            Busca_Grati = 1
       Case Is = 12
            Sql$ = "select * from plahistorico where cia='" & pCodCia & "' and placod='" & pCodTrab & "' and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 0 Else Busca_Grati = 1
End Select

If rsGrati.State = 1 Then rsGrati.Close
Set rsGrati = Nothing

End Function

'APORTACIONES
Public Function F03(ByVal pCodConcepto As String, ByVal pCodCia As String, ByVal pCodArea As String, ByVal pTipoBol As String, ByVal pCodTrab As String, _
ByVal Paño As String, ByVal pMes As String, ByVal pStatus As String, pTipoTrabajor As String) As String  'APORTACIONES


Dim rsF03 As ADODB.Recordset
Dim rscalculo As ADODB.Recordset
Dim F03str As String
Dim mFactor As Currency

Dim mcancel As Boolean
Dim wciamae As String

Dim macui As Currency
Dim macus As Currency


mFactor = 0
F03 = ""
If pCodConcepto = "03" Then
   Sql$ = "select senati from cia where cod_cia='" & pCodCia & "' and status<>'*'"
   If Not (fAbrRst(rsF03, Sql$)) Then Exit Function
   If Trim(rsF03!senati) <> "S" Then Exit Function
   If rsF03.State = 1 Then rsF03.Close
   wciamae = Determina_Maestro("01044")
   Sql$ = "Select * from maestros_2 where cod_maestro2='" & Trim(pCodArea) & "' and status<>'*'"
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rsF03, Sql$)) Then
      If rsF03!flag7 <> "S" Then Exit Function
   Else
      Exit Function
   End If
   If rsF03.State = 1 Then rsF03.Close
End If


If pCodConcepto = "18" Or pCodConcepto = 19 Then
 Dim mperiodosctr, xcodsctr As String
 Dim xporsalud, xporpension As Currency
 Dim xtopemax As Double
 Dim mtope As Double
 
 mperiodosctr = ""
 
 
 Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='A' and tboleta='" & pTipoBol & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"

 If (fAbrRst(rsF03, Sql$)) Then
    rsF03.MoveFirst
    F03str = ""
    Do While Not rsF03.EOF
        F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
          
        rsF03.MoveNext
    Loop
       
    F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
   
    mperiodosctr = Format(Paño, "0000") & Format(pMes, "00")
    
   'mperiodosctr = Format(Vano, "0000") & Format(Vmes, "00")
    
    xcodsctr = Bcodsctr(pCodTrab, pCodCia)
    
    Sql$ = "SELECT porsalud,porpension,tope FROM PLASCTR where periodo='" & mperiodosctr & "' and codSCTR='" & xcodsctr & "' and status<>'*' and cia='" & pCodCia & "'"
    
       If Not (fAbrRst(rsF03, Sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para SCTR", vbCritical, "Calculo de Boleta"
          mcancel = True
          Exit Function
       Else
        xporsalud = rsF03!porsalud
        xporpension = rsF03!porpension
        xtopemax = rsF03!tope
       End If
    
    If pCodConcepto = "18" Then
     mFactor = xporsalud
     
    Else
     mFactor = xporpension
    End If
    mtope = xtopemax
     
       
  'CALL Acumula_Mes_SCTR(concepto, "A", )
  
  'IMPLEMENTACION GALLOS
  
   Call Acumula_Mes_SCTR(pCodConcepto, "A", pCodCia, pCodTrab, Paño, pMes, macui, macus)
   
   F03 = "select (" & F03str & " + " & macui & ") , (" & mFactor & " /100)," & macus & " from platemphist "
      F03 = F03 & "where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' "
      F03 = F03 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
      
      
      
      If (fAbrRst(rscalculo, F03)) Then
        
        If rscalculo(0) > xtopemax And pCodConcepto = "19" Then
                F03 = " SELECT " & Round((mtope * (mFactor / 100)) - macus, 2)
        Else
                F03 = " SELECT " & Round((rscalculo(0) * (mFactor / 100)) - macus, 2)
        End If
        
      End If
      
 End If
 macui = 0: macus = 0
 Exit Function
End If

    'MODIFICACION PARA EL CALCULO DE LA EPS
    Dim rsTemporal  As ADODB.Recordset
    Dim Afecto_EPS  As Boolean
    Dim Por_EPS     As Double
    Cadena = "SELECT ISNULL(APORTACION_EPS,0) AS APORTACION_EPS FROM CIA WHERE COD_CIA = '" & pCodCia & "' AND STATUS != '*'"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    If Not rsTemporal.EOF Then
        Afecto_EPS = CBool(rsTemporal!Aportacion_EPS)
    Else
        Afecto_EPS = False
    End If
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    If Afecto_EPS Then
        Cadena = "SELECT ISNULL(CODIGO_EPS,'') AS CODIGO_EPS FROM PLANILLAS WHERE CIA = '" & pCodCia & "' AND PLACOD = '" & pCodTrab & "' AND STATUS != '*' AND TIPOTRABAJADOR = '" & pTipoTrabajor & "'"
        Set rsTemporal = OpenRecordset(Cadena, cn)
        If Not rsTemporal.EOF Then
            If rsTemporal!codigo_eps <> "" Then
                Cadena = "SELECT ISNULL(IMPORTE,0) AS IMPORTE FROM MAESTROS_2 WHERE RIGHT(CIAMAESTRO,3) = '143' AND STATUS != '*' AND RTRIM(ISNULL(CODSUNAT,'')) != '' AND COD_MAESTRO2 = '" & Trim(rsTemporal!codigo_eps) & "'"
                If rsTemporal.State = adStateOpen Then rsTemporal.Close
                Set rsTemporal = OpenRecordset(Cadena, cn)
                If Not rsTemporal.EOF Then
                    Por_EPS = Val(rsTemporal!importe)
                Else
                    Por_EPS = 0
                End If
            Else
                Por_EPS = 0
            End If
        Else
            Por_EPS = 0
        End If
    End If
    
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    
    Sql$ = "select aportacion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & pCodConcepto & "' and aportacion<>0 and status<>'*'"
    If (fAbrRst(rsF03, Sql$)) Then
       If Not IsNull(rsF03!aportacion) Then
          If rsF03!aportacion <> 0 Then mFactor = IIf(Afecto_EPS = True, Por_EPS, rsF03!aportacion)
       End If
    End If
    If rsF03.State = 1 Then rsF03.Close
    

If mFactor <> 0 Then
   Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='A' and tboleta='" & pTipoBol & "'  and  codigo='" & pCodConcepto & "' and status<>'*'"
   
   If (fAbrRst(rsF03, Sql$)) Then
      rsF03.MoveFirst
      F03str = ""
      Do While Not rsF03.EOF
         F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
         rsF03.MoveNext
      Loop
      
      F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
      If rsF03.State = 1 Then rsF03.Close
      
           
      Call Acumula_Mes(pCodCia, pCodTrab, Paño, pMes, pCodConcepto, "A", macui, macus, pTipoBol)
      
      F03 = "select (" & F03str & " + " & macui & ") , (" & mFactor & " /100)," & macus & " from platemphist "
      F03 = F03 & "where cia='" & pCodCia & "' and status='" & pStatus & "' and placod='" & pCodTrab & "' "
      F03 = F03 & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & ""
      
      
      If (fAbrRst(rscalculo, F03)) Then

        If pCodConcepto = "01" Then
        
            If Mid(pCodTrab, 2, 1) = "O" Then
                    F03 = " SELECT " & Round(((rscalculo(0) * mFactor / 100) - macus), 2)
            Else
        
                If rscalculo(0) > sueldominimo Then
                    F03 = " SELECT " & Round(((rscalculo(0) * mFactor / 100) - macus), 2)
                Else
                    F03 = " SELECT " & Round(((rscalculo(0) * mFactor / 100) - macus), 2)
                End If
            End If
        Else
            F03 = " SELECT " & Round(((rscalculo(0) * mFactor / 100) - macus), 2)
        End If

        
      End If
   End If
End If

macui = 0: macus = 0
End Function

Private Function Descarga_ctaCte(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal pCodAuxInt As String, ByVal pTipoBol As String, ByVal pFechaProc As String, ByVal pNroSemana As String, ByVal pTipoTrab As String, ByVal pImporte As Currency) As Boolean
On Error GoTo ErrMsg:

Dim Sql As String
Dim RX As ADODB.Recordset
Dim Quincena As Integer
'SQL = "select * from plactacte where cia='" & pCodCia & "' and placod='" & Trim(pCodTrab) & "' and importe-pago_acuenta>0 and status<>'*' order by fecha"
'If (fAbrRst(rs, SQL$)) Then rs.MoveFirst
'Do While Not rs.EOF
'   saldo = (rs!importe - rs!pago_acuenta)
'   If saldo >= importe Then
'      If saldo = importe Then
'         'SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & pCodCia & "' and numinterno='" & RS!numinterno & "' and tipo='" & RS!tipo & "' and status<>'*'"
'         SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & pCodCia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'         cn.Execute SQL
'      Else
'         'SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & pCodCia & "' and numinterno='" & RS!numinterno & "' and tipo='" & RS!tipo & "' and status<>'*'"
'         SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & pCodCia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'         cn.Execute SQL
'      End If

      Dim I As Integer
      Quincena = 0
      MAXROW = UBound(ArrDsctoCTACTE, 2) + 1
      For I = 0 To MAXROW - 1
        If ArrDsctoCTACTE(0, I) = Trim(pCodTrab) Then
            If ArrDsctoCTACTE(2, I) > 0 Then
            
                Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                Sql = Sql & "insert into plabolcte values('" & pCodCia & "','" & UCase(Trim(pCodTrab)) & "','" & pTipoBol & "','', " _
                & "'" & pFechaProc & "','" & pNroSemana & "','" & pTipoTrab & "','" & Trim(pCodAuxInt) & "','" & ArrDsctoCTACTE(3, I) & "','" & pFechaProc & "', " _
                & "'" & wmoncont & "'," & ArrDsctoCTACTE(2, I) & ",'','" & wuser & "'," & FechaSys & "," & ArrDsctoCTACTE(1, I) & "," & IIf(wtipodoc = True, 0, 1) & ")"
                
                cn.Execute Sql
            End If
        
            Sql = "UPDATE plactacte set pago_acuenta=pago_acuenta + " & CStr(ArrDsctoCTACTE(2, I)) + " WHERE cia='" & pCodCia & "' and placod='" & Trim(pCodTrab) & "' and id_doc='" & CStr(ArrDsctoCTACTE(1, I)) & "' and status<>'*'"
            cn.Execute Sql
            
            Sql = "UPDATE plactacte set fecha_cancela='" & pFechaProc & "' WHERE cia='" & pCodCia & "' and placod='" & Trim(pCodTrab) & "' and id_doc='" & ArrDsctoCTACTE(1, I) & "' and status<>'*' AND IMPORTE=PAGO_ACUENTA"
            cn.Execute Sql
                
        End If
      Next
      'Exit Do
'   Else
'      SQL = "update plactacte set pago_acuenta=pago_acuenta+" & saldo & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & pCodCia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'      cn.Execute SQL
'
'      SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
'      SQL = SQL & "insert into plabolcte values('" & pCodCia & "','" & pCodTrab & "','" & pTipoBol & "','', " _
'          & "'" & Format(pFechaProc, FormatFecha) & "','" & pNroSemana & "','" & pTipoTrab & "','" & rs!codauxinterno & "','" & rs!tipo & "','" & Format(rs!fecha, FormatFecha) & "', " _
'          & "'" & wmoncont & "'," & saldo & ",'','" & wuser & "'," & FechaSys & ")"
'      cn.Execute SQL
'      importe = importe - saldo
'   End If
'   rs.MoveNext
'Loop
'rs.Close
Descarga_ctaCte = True
Exit Function
ErrMsg:
Descarga_ctaCte = False
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, "Descarga_ctaCte"
End Function
Public Function EliminaBoleta(ByVal pCodCia As String, ByVal pPagoQuincena As Boolean, ByVal Paño As Integer, ByVal pMes As Integer, ByVal pDia As Integer, ByVal pCodTrab As String, ByVal VPerPago As String, _
            ByVal pTipoBol As String, ByVal pNroSemana As String, ByVal pRecalculoBoleta As Boolean, ByVal pNumQuincena As String) As Boolean

On Error GoTo ErrMsg:

Sql$ = ""

EliminaBoleta = False

cn.Execute "begin transaction", 64

 
 Select Case VPerPago
        Case Is = "02"
              'PLANILLA OBREROS
             Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
             & "where cia='" & pCodCia & "' and proceso='" & pTipoBol & "' and semana='" & pNroSemana & "' and year(fechaproceso)=" & Paño & " " _
             & "and status<>'*' and placod='" & pCodTrab & "'"
        Case Is = "04"
              'PLANILLA SUELDOS
             Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
             & "where cia='" & pCodCia & "' and proceso='" & pTipoBol & "' and Year(fechaproceso) = " & Paño & " And month(fechaproceso) = " & pMes & " " _
             & "and status<>'*' and placod='" & pCodTrab & "'"
End Select

If Sql <> "" Then
    cn.Execute Sql$
End If

If pNumQuincena = "1" Then
   'PRIMERA QUINCENA
   
   Sql$ = "update plaquincena set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
   & "where cia='" & pCodCia & "' and Year(fechaproceso) = " & Paño & " And month( fechaproceso) = " & pMes & " " _
   & "and status<>'*' and placod='" & pCodTrab & "'"
End If


cn.Execute Sql$


If pRecalculoBoleta = False Then

   Sql$ = "select * from plabolcte " _
   & "where cia='" & pCodCia & "' and placod='" & UCase(Trim(pCodTrab)) & "' and proceso='" & pTipoBol & "' and semana='" & pNroSemana & "' and year(fechaproceso)=" & Paño & " " _
   & "and month( fechaproceso) = " & pMes & " and day( fechaproceso) = " & pDia & " and status<>'*'"
   If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
   Do While Not rs.EOF
      Sql = "update plactacte set fecha_cancela=null,pago_acuenta=pago_acuenta-" & rs!importe & " where cia='" & pCodCia & "' and placod='" & UCase(Trim(pCodTrab)) & "' and tipo='" & rs!tipo & "' and id_doc=" & rs!id_doc
      cn.Execute Sql$
      rs.MoveNext
   Loop
   Sql$ = "update plabolcte set status='*'" _
   & "where cia='" & pCodCia & "' and placod='" & UCase(Trim(pCodTrab)) & "' and proceso='" & pTipoBol & "' and semana='" & pNroSemana & "' and year(fechaproceso)=" & Paño & " " _
   & "and month( fechaproceso) = " & pMes & " and day( fechaproceso) = " & pDia & " and status<>'*'"
   cn.Execute Sql$
   
End If


cn.Execute "commit transaction", 64
EliminaBoleta = True
Exit Function
ErrMsg:
EliminaBoleta = False
cn.Execute "rollback transaction", 64
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, ""
End Function
Public Function Traer_SueldoBasico_Trabajador(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal pCodConcepto As String, Optional pError As Boolean) As Currency
pError = False
Sql = "select importe from plaremunbase where cia='" & pCodCia & "' and placod='" & Trim(pCodTrab) & "'"
Sql = Sql & " and concepto='" & Trim(pCodConcepto) & "' and status<>'*'"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
    Traer_SueldoBasico_Trabajador = IIf(IsNull(Rq(0)), 0, Rq(0))
Else
    Traer_SueldoBasico_Trabajador = 0
    MsgBox "No existe Ssueldo Básico del Trabajador", vbCritical, "Traer_SueldoBasico_Trabajador"
    pError = True
End If
Rq.Close
Set Rq = Nothing
End Function

Public Function Traer_TipoTablaMaestra(ByVal pCodCia As String, ByVal pCodMaestros As String, Optional pError As Boolean) As String
'//***** Devuelve el criterio de busqueda del maestro segun su seteo si es General para todas las Cias
'//***** o si es particular por cada cia

Sql = "SELECT GENERAL FROM maestros where "
Sql = Sql & " right(ciamaestro,3)='" & Right(pCodMaestros, 3) & "' and status<>'*' "
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
   If Rq!General = "S" Then
      Traer_TipoTablaMaestra = " and right(ciamaestro,3)= '" & Right(pCodMaestros, 3) & "'"
   Else
      Traer_TipoTablaMaestra = " and ciamaestro= '" & pCodCia + Right(pCodMaestros, 3) & "'"
   End If
End If
Rq.Close
Set Rq = Nothing
End Function

Public Function Cantidad_ymd_Trasncurridos(ByRef pFechaProceso, ByRef pFechaNac, pTipo, Optional pError As Boolean) As Long
'//** Funcion Cantidad_ymd_Trasncurridos apartir de 2 fechas  **//
'//** pTipo: a=años, m=meses, d=dias

'antes Public Function perendat(fecha1, fecha2, tipo)
'manos = perendat(VFProceso, VFechaNac, "a")
' tipo: a=años, m=meses, d=dias

On Error GoTo MsgError:
Dim ca As Long, cm As Long, cd As Long
Dim f1 As Variant, f2 As Variant

If CDate(pFechaProceso) < CDate(pFechaNac) Then
    f1 = pFechaProceso: f2 = pFechaNac
Else
    f2 = pFechaProceso: f1 = pFechaNac
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
Select Case pTipo
Case "d"
    Cantidad_ymd_Trasncurridos = cd
Case "m"
    Cantidad_ymd_Trasncurridos = cm
Case "a"
    Cantidad_ymd_Trasncurridos = ca
End Select

Exit Function
MsgError:
pError = True
MsgBox ERR.Number & " - " & ERR.Description
End Function

Public Function CargaDatos_RecalculoBoleta(ByVal pCodCia As String, ByVal pCodTrab As String, ByVal pTipoBol As String, ByVal pNroSemana As String, ByVal Paño As String, ByVal pMes As String, ByVal pPagoQuincena As Boolean, ByVal pTipoPeriodoPago As String, _
ByRef pRsHoras As ADODB.Recordset, ByRef pRsCosto As ADODB.Recordset, ByRef pRsPagAdic As ADODB.Recordset, ByRef pRsDesAdic As ADODB.Recordset) As Boolean


If pPagoQuincena = True Then
      Select Case pTipoPeriodoPago
             Case Is = "02"
                'PLANILLA DE OBREROS
                  Sql = "select * from plahistorico "
                  Sql = Sql & " where cia='" & pCodCia & "' and proceso='" & pTipoBol & "' and semana='" & pNroSemana & "' and year(fechaproceso)=" & Paño & " "
                  Sql = Sql & " and status<>'*' and placod='" & Trim(pCodTrab) & "'"
             Case Is = "04"
                'PLANILLA DE EMPLEADOS
                  Sql = "select * from plahistorico "
                  Sql = Sql & " where cia='" & pCodCia & "' and proceso='" & pTipoBol & "' and year( fechaproceso) = " & Paño & " And Month(fechaproceso) = " & pMes & " "
                  Sql = Sql & " and status<>'*' and placod='" & Trim(pCodTrab) & "'"
      End Select
   Else
      Sql = "select * from plaquincena "
      Sql = Sql & " where cia='" & pCodCia & "' and year( fechaproceso) = " & Paño & " And Month(fechaproceso) = " & pMes & " "
      Sql = Sql & " and status<>'*' and placod='" & Trim(pCodTrab) & "'"
   End If
   Dim Rq As ADODB.Recordset
    
   LimpiarRs pRsHoras
   LimpiarRs pRsCosto
   LimpiarRs pRsPagAdic
   LimpiarRs pRsDesAdic
   
   If fAbrRst(Rq, Sql) Then
   Dim intloop  As Integer
    Do While Not Rq.EOF
                intloop = 0
                With Rq
                        Dim xValue As Variant
                        For intloop = 0 To .Fields.count - 1
                             If Left(.Fields(intloop).Name, 1) = "h" Then
                                    pRsHoras.AddNew
                                    pRsHoras!codigo = Right(.Fields(intloop).Name, 2)
                                    pRsHoras!Monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)
                             End If
                             
                             If Left(.Fields(intloop).Name, 3) = "cco" Or Left(.Fields(intloop).Name, 3) = "por" Then
                                  
                                   pRsCosto.FIND "item='" & Trim(Right(.Fields(intloop).Name, 1) & "") & "'", 0, 1, 1
                                    If pRsCosto.EOF Then
                                        pRsCosto.AddNew
                                        pRsCosto!Item = Trim(Right(.Fields(intloop).Name, 1))
                                        If Left(.Fields(intloop).Name, 3) = "cco" Then
                                            pRsCosto!codigo = Trim(IIf(Trim(.Fields(intloop).Value & "") = "", 0, Format(.Fields(intloop).Value, "00")))
                                        Else
                                            pRsCosto!codigo = ""
                                        End If
                                        If Left(.Fields(intloop).Name, 3) = "por" Then
                                            pRsCosto!Monto = Trim(IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value))
                                        Else
                                            pRsCosto!Monto = 0
                                        End If
                                    Else
                                        If Left(.Fields(intloop).Name, 3) = "cco" Then
                                            pRsCosto!codigo = Trim(IIf(Trim(.Fields(intloop).Value & "") = "", 0, Format(.Fields(intloop).Value, "00")))
                                        End If
                                        If Left(.Fields(intloop).Name, 3) = "por" Then
                                            pRsCosto!Monto = Trim(IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value))
                                        End If
                                    End If
                             
                             End If
                             
                             If Left(.Fields(intloop).Name, 1) = "i" Then
                                    pRsPagAdic.AddNew
                                    pRsPagAdic!codigo = Right(.Fields(intloop).Name, 2)
                                    pRsPagAdic!Monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)
                             End If
                             
                              If Left(.Fields(intloop).Name, 1) = "d" And Len(.Fields(intloop).Name) = 3 Then
                                    pRsDesAdic.AddNew
                                    If Len(.Fields(intloop).Name) = 3 Then
                                        pRsDesAdic!codigo = Right(.Fields(intloop).Name, 2)
                                    Else
                                        pRsDesAdic!codigo = Right(.Fields(intloop).Name, 3)
                                    End If
                                    pRsDesAdic!Monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)
                             End If
                             
                        Next
                End With
                Exit Do
        Rq.MoveNext
    Loop
      
'      Dim MField As String
'      With pRsHoras
'         If .RecordCount > 0 Then .MoveFirst
'         Do While Not .EOF
'                If Left(Rq.Fields(), 1) = "h" Then
'                    .AddNew
'                    MField = "h" & !codigo
'                    !Monto = Rq.Fields(MField)
'                End If
'             .MoveNext
'         Loop
'      End With
'
'      With pRsCosto
'         If .RecordCount > 0 Then .MoveFirst
'         Do While Not .EOF
'                MField = "ccosto" & Right(Rq.Fields(MField), 1)
'                !Monto = Rq.Fields(MField)
'             .MoveNext
'         Loop
'      End With
'
'      With pRsPagAdic
'         If .RecordCount > 0 Then .MoveFirst
'         Do While Not .EOF
'           MField = "i" & !codigo
'           !Monto = Rq.Fields(MField)
'           .MoveNext
'         Loop
'      End With
'
'      With pRsDesAdic
'         If .RecordCount > 0 Then .MoveFirst
'         Do While Not .EOF
'            MField = "d" & !codigo
'            !Monto = Rs.Fields(MField)
'            .MoveNext
'         Loop
'      End With
   
   End If
   
   Rq.Close
   Set Rq = Nothing

End Function

Private Sub LimpiarRs(ByRef pRs As ADODB.Recordset)
If pRs.State = 1 Then
    If pRs.RecordCount > 0 Then
        pRs.MoveFirst
        Do While Not pRs.EOF
            pRs.Delete
            pRs.MoveNext
        Loop
    End If
End If
End Sub



Private Function Calcular_Dscto_CtaCte_Boleta(ByRef pRsDesAdic As ADODB.Recordset, ByVal pCodCia As String, ByVal pCodTrab As String, ByVal pTipoBol As String, ByVal pFechaProceso As String, ByVal pPagoQuincena As Boolean)
Dim SwError As Boolean
Dim RX As New ADODB.Recordset
Dim I As Integer
Dim sumporc As Currency
On Error GoTo CORRIGE
Dim MAXROW As Integer
MAXROW = 0
SwError = True
    Dim Sql As String
    If pTipoBol = "01" Then 'normal
        Sql = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(isnull(b.importe,0)),0) as pago_acuenta,a.partes,"
        Sql = Sql & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
        Sql = Sql & " AND MONTH(fechaproceso)=" & Month(CDate(pFechaProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(pFechaProceso)) & ") as saldomes,"
        Sql = Sql & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
        Sql = Sql & " AND MONTH(fechaproceso)=" & Month(CDate(pFechaProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(pFechaProceso)) & ") as quincena"
        Sql = Sql & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
        Sql = Sql & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & pCodCia & "' and a.placod='" & Trim(pCodTrab) & "' AND a.STATUS<>'*'"
        Sql = Sql & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
    Else 'gratificacion
        Sql = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(isnull(b.importe),0) as pago_acuenta,a.partes,"
        Sql = Sql & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
        Sql = Sql & " AND MONTH(fechaproceso)=" & Month(CDate(pFechaProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(pFechaProceso)) & ") as saldomes,"
        Sql = Sql & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
        Sql = Sql & " AND MONTH(fechaproceso)=" & Month(CDate(pFechaProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(pFechaProceso)) & ") as quincena"
        Sql = Sql & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
        Sql = Sql & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & pCodCia & "' and a.placod='" & Trim(pCodTrab) & "' AND a.STATUS<>'*' and a.sn_grati=1"
        Sql = Sql & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
    End If

   If fAbrRst(RX, Sql) Then

      If Not pRsDesAdic.EOF Then pRsDesAdic.MoveFirst
      Erase ArrDsctoCTACTE
      MAXROW = 0
      Do While Not pRsDesAdic.EOF
         If Trim(pRsDesAdic("CODIGO")) = "07" Then 'ctacte
            Do While Not RX.EOF
                ReDim Preserve ArrDsctoCTACTE(0 To 4, 0 To MAXROW)
                
                ArrDsctoCTACTE(0, MAXROW) = Trim(pCodTrab)
                ArrDsctoCTACTE(1, MAXROW) = RX!id_doc
                ArrDsctoCTACTE(3, MAXROW) = RX!tipo
                ArrDsctoCTACTE(4, MAXROW) = 0
                
                If RX("partes") >= (RX("IMPORTE") - RX("PAGO_ACUENTA")) Then
                   pRsDesAdic("MONTO") = pRsDesAdic("MONTO") + (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                   If pPagoQuincena = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("IMPORTE") - RX("PAGO_ACUENTA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                Else
                    'ojo revisar
                   pRsDesAdic("MONTO") = pRsDesAdic("MONTO") + IIf(RX("partes") = 0, 0, RX("partes") - RX("QUINCENA"))
                   If pPagoQuincena = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("partes") - RX("QUINCENA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = RX("partes") - RX("QUINCENA")
                End If
                
'                If pPagoQuincena = False Then
'                    If pRsDesAdic("MONTO") > 0 Then
'                        pRsDesAdic("MONTO") = Round(pRsDesAdic("MONTO") / 2, 2)
'                    End If
'                End If
                
                MAXROW = MAXROW + 1
                
                RX.MoveNext
            Loop
            
             If pPagoQuincena = False Then
                If pRsDesAdic("MONTO") > 0 Then
                    pRsDesAdic("MONTO") = Round(pRsDesAdic("MONTO") / 2, 2)
                End If
            End If
                
            sumporc = 0
            If pRsDesAdic("MONTO") > 0 Then
                For I = 0 To MAXROW - 1
                
                    ArrDsctoCTACTE(4, I) = Round((ArrDsctoCTACTE(2, I) * 100) / pRsDesAdic("MONTO"), 2)
                    sumporc = sumporc + Round((ArrDsctoCTACTE(2, I) * 100) / pRsDesAdic("MONTO"), 2)
                    If I = MAXROW - 1 Then
                        If sumporc > 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) - (sumporc - 100)
                        ElseIf sumporc < 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) + (100 - sumporc)
                        End If
                    End If
                Next
            End If
         End If
         pRsDesAdic.MoveNext
      Loop
   End If

RX.Close
Set RX = Nothing

Calcular_Dscto_CtaCte_Boleta = SwError
Exit Function
CORRIGE:
SwError = False
Calcular_Dscto_CtaCte_Boleta = SwError
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, "Calcular_Dscto_CtaCte_Boleta"
End Function

Private Sub Crea_Rs(ByRef pRsHoras As ADODB.Recordset, ByRef pRsCosto As ADODB.Recordset, ByRef pRsPagAdic As ADODB.Recordset, ByRef pRsDesAdic As ADODB.Recordset)
    Set pRsHoras = New ADODB.Recordset
    If pRsHoras.State = 1 Then pRsHoras.Close
    pRsHoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
    pRsHoras.Fields.Append "descripcion", adChar, 100, adFldIsNullable
    pRsHoras.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    pRsHoras.Open
    
    Set pRsCosto = New ADODB.Recordset
    If pRsCosto.State = 1 Then pRsCosto.Close
    pRsCosto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    pRsCosto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    pRsCosto.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    pRsCosto.Fields.Append "item", adChar, 2, adFldIsNullable
    pRsCosto.Open
    
    
    Set pRsPagAdic = New ADODB.Recordset
    If pRsPagAdic.State = 1 Then pRsPagAdic.Close
    pRsPagAdic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    pRsPagAdic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    pRsPagAdic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    pRsPagAdic.Open
    
    Set pRsDesAdic = New ADODB.Recordset
    If pRsDesAdic.State = 1 Then pRsDesAdic.Close
    pRsDesAdic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    pRsDesAdic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    pRsDesAdic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    pRsDesAdic.Open
    
End Sub

'IMPLEMENTACION GALLOS

Sub Acumula_Mes_SCTR(ByVal pCodConcepto As String, ByVal tipo As String, ByVal pCodCia As String, ByVal pCodTrab As String, ByVal Paño As String, ByVal pMes As String, ByRef Pmacui As Currency, ByRef Pmacus As Currency)

Dim mcad As String
Dim SqlAcu As String
Dim rs2 As New Recordset
Dim RsAcumula As New Recordset

Pmacui = 0: Pmacus = 0

    Sql$ = "select cod_remu from plaafectos where cia='" & pCodCia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & pCodConcepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       
       SqlAcu = "select sum(" & mcad & ") as ing,sum(a" & pCodConcepto & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & pCodCia & "' and status<>'*' and placod='" & pCodTrab & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Paño & " and month(fechaproceso)=" & pMes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','04','05')"
       'SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
       
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then Pmacui = Pmacui + rs2!ing
          If Not IsNull(rs2!ded) Then Pmacus = Pmacus + rs2!ded
       End If
    End If

If rs2.State = 1 Then rs2.Close
Set rs2 = Nothing

If RsAcumula.State = 1 Then RsAcumula.Close
Set RsAcumula = Nothing

End Sub

Private Function Bcodsctr(pCodTrab As String, pCodCia As String)
Dim Sql As String
Dim RZ As New ADODB.Recordset

    Sql = "select p.codsctr " & _
    " from planillas p " & _
    " where p.status<>'*'  " & _
    " and p.cia='" & pCodCia & "' and p.placod='" & pCodTrab & "' "
    
    RZ.Open Sql, cn, adOpenStatic, adLockReadOnly
    
    If RZ.RecordCount = 0 Then
      Bcodsctr = 0
    Else
      Bcodsctr = RZ!codsctr
    End If
RZ.Close
End Function


