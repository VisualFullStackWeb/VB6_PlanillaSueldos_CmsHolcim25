VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMGenera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Generar Asientos «"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "FrmMGenera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4635
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   4410
      Begin VB.ComboBox CmbBoleta 
         Height          =   315
         ItemData        =   "FrmMGenera.frx":030A
         Left            =   1410
         List            =   "FrmMGenera.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1350
         Width           =   2865
      End
      Begin VB.ComboBox CmbTrabTipo 
         Height          =   315
         ItemData        =   "FrmMGenera.frx":030E
         Left            =   1410
         List            =   "FrmMGenera.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   975
         Width           =   2865
      End
      Begin VB.ComboBox CmbMes 
         Height          =   315
         ItemData        =   "FrmMGenera.frx":0312
         Left            =   1410
         List            =   "FrmMGenera.frx":033A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2865
      End
      Begin VB.TextBox TxtAno 
         Height          =   285
         Left            =   1410
         MaxLength       =   4
         TabIndex        =   3
         Top             =   225
         Width           =   915
      End
      Begin MSComctlLib.ProgressBar P1 
         Height          =   165
         Left            =   75
         TabIndex        =   9
         Top             =   1800
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Boleta"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   975
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes Proceso"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año Proceso"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmMGenera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_MGenera As ADODB.Recordset
Dim rs_MGenera2 As ADODB.Recordset
Dim rs_MGenera3 As ADODB.Recordset
Dim rs_MGenera_Boletas As ADODB.Recordset
'Dim s_CodTipo As String
Dim s_MesSeleccion As String
Dim s_Asiento_Neto As String
Dim s_Envio_Clave_Boleta As String
Dim i_Monto_Ipss As Double
Dim Suma_Cuenta_79 As Double
Dim Suma_Cuenta_79_negativo As Double
Dim s_CtaIpss As String
Dim i_Numero_Voucher As Integer
Dim i_Numero_VoucherG As String
Dim s_Tipo_Boleta As String
Dim s_TipoTrabajador As String
Dim s_TipoTrabajador_Report As String
Dim s_CentroCosto As String
Dim mCenCosto As String
Dim MArea As String
Dim s_CodEmpresa_Starsoft As String
'***************variables utilizadas para reportes de contabilidad***************************
Public s_Año_ProcesoReport As String
Public s_Mes_ProcesoReport As String
Public s_Tip_TrabajadorReport As String
Public s_Tip_BoletaReport As String
Public i_TipoReporte As Integer

Private Const Concepto_Utilidad = "i18"
Private Const Cuenta_dUtilidades = "41300000"

Dim I               As Integer
Dim Centro_Costo    As String
Dim Por_CC          As Double
Dim dImporte        As Double
'Dim dDiferencia     As Double
'Dim dTotal          As Double
Dim Array_CC()      As Variant
Dim Array_N()       As Variant
Dim Array_A()       As Variant
Dim Array_D()       As Variant
Dim Array_I()       As Variant
'Dim Array_79()      As Variant
Dim bBoleano        As Boolean
'Dim rsTemporal      As ADODB.Recordset

Private Sub Arma_Arreglo(mArray() As Variant, mRecorset As ADODB.Recordset)
    Erase mArray
    If Not mRecorset.EOF Then
        With mRecorset
            .MoveFirst
            ReDim mArray(1 To 3, 1 To .RecordCount)
            Do While Not .EOF
                mArray(1, .AbsolutePosition) = !Clave 'IDENTIFICADOR DEL CONCEPTO
                mArray(2, .AbsolutePosition) = !total_montos 'IMPORTE
                mArray(3, .AbsolutePosition) = !total_montos 'DIFERENCIA
                .MoveNext
            Loop
            .MoveFirst
        End With
    End If
End Sub

Private Function Calculo_dImporte(mArray() As Variant) As Double
    Calculo_dImporte = 0: Calculo_dImporte = Round((Val(mArray(2, rs_MGenera.AbsolutePosition)) * Array_CC(I)) / 100, 2)
    If UBound(Array_CC) = I Then
        Calculo_dImporte = Val(mArray(3, rs_MGenera.AbsolutePosition))
    Else
        mArray(3, rs_MGenera.AbsolutePosition) = Val(mArray(3, rs_MGenera.AbsolutePosition)) - Calculo_dImporte
    End If
End Function

Private Sub Crea_Array_CC()
    Erase Array_CC
    For I = 1 To 5
        Centro_Costo = Empty: Por_CC = 0: dImporte = 0
        Centro_Costo = Trim(rs_MGenera_Boletas.Fields.Item("CCOSTO" & I))
        Por_CC = rs_MGenera_Boletas.Fields.Item("PORC" & I)
        If Trim(Centro_Costo) = "" Then Exit For
        ReDim Preserve Array_CC(1 To I)
        Array_CC(I) = Por_CC
    Next
End Sub



Private Sub CmbTrabTipo_Click()
'    Call Recupera_Codigo_Tipo_Trabajador(CmbTrabTipo.Text)
'    s_CodTipo = Crear_Plan_Contable.s_CodTipo_G
'    s_CodTipo = Trim(fc_CodigoComboBox(CboTipo_Trab, 2))
    Call Recupera_tipo_trabajador
End Sub
Sub Llena_Tipo_Trabajadores()
    Call Trae_Tipo_Trab(CmbTrabTipo)
End Sub
Private Sub Form_Activate()
    Call Llena_Tipo_Trabajadores
    Call Llena_Barra
    Call Trae_Tipo_Boleta(CmbBoleta, 2)
    CmbBoleta.AddItem "EXTORNO VACACIONES"
End Sub
Sub Captura_Mes_Seleccionado()
    Select Case CmbMes.Text
        Case "ENERO": s_MesSeleccion = "01": Case "FEBRERO": s_MesSeleccion = "02"
        Case "MARZO": s_MesSeleccion = "03": Case "ABRIL": s_MesSeleccion = "04"
        Case "MAYO": s_MesSeleccion = "05": Case "JUNIO": s_MesSeleccion = "06"
        Case "JULIO": s_MesSeleccion = "07": Case "AGOSTO": s_MesSeleccion = "08"
        Case "SETIEMBRE": s_MesSeleccion = "09": Case "OCTUBRE": s_MesSeleccion = "10"
        Case "NOVIEMBRE": s_MesSeleccion = "11": Case "DICIEMBRE": s_MesSeleccion = "12"
    End Select
End Sub
Sub Proceso_Ingresos_Asiento_Naturaleza_Obreros(Opcion As Integer)
    Dim i_Contador As Integer
    For i_Contador = 1 To 2
        Call Recupera_Codigos_Conceptos_Ingresos("I", s_TipoTrabajador, wcia)
        Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
        If bBoleano = False Then Call Arma_Arreglo(Array_I, rs_MGenera)
        Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
        If rs_MGenera.EOF = False Then
            rs_MGenera.MoveFirst
            Do While Not rs_MGenera.EOF
                Debug.Print rs_MGenera!Clave
                                              
                
                Select Case i_Contador
                    Case 1
                       Call Graba_Informacion_Total(Trim(rs_MGenera!Clave), rs_MGenera_Boletas!FechaProceso, Trim(s_Tipo_Boleta), Trim(s_TipoTrabajador), Trim(rs_MGenera_Boletas!PlaCod), Trim(rs_MGenera_Boletas!semana), wcia)
'                       Call Graba_Informacion_Totales(rs_MGenera!Clave, txtano, s_MesSeleccion, _
'                        s_Tipo_Boleta, s_TipoTrabajador, rs_MGenera_Boletas!PLACOD, rs_MGenera_Boletas!semana, wcia)
                        
                        rs_MGenera.MoveNext
                    Case 2
                        Select Case Opcion
                            Case 1
                                Call Conectar_Base_Datos_Access(s_CodEmpresa_Starsoft)
                                 If Len(rs_MGenera!cma_ctacontable) > 4 Then
                                    If Len(rs_MGenera!cma_ctacontable) > 5 Then
                                        Call Recuperar_Cuentas_Naturaleza_Transferencia(rs_MGenera!cma_ctacontable)
                                    Else
                                        Call Recuperar_Cuentas_Naturaleza_Transferencia(s_CentroCosto & rs_MGenera!cma_ctacontable)
                                    End If
                                 Else
                                    Call Recuperar_Cuentas_Naturaleza_Transferencia(s_CentroCosto & rs_MGenera!cma_ctacontable)
                                 End If
                                
                                
                                Set rs_MGenera2 = CompRetenciones.rs_compRetenciones_Pub
                                Set CompRetenciones.rs_compRetenciones_Pub = Nothing
                                If rs_MGenera2.EOF = False Then
                                    If Trim(rs_MGenera!Clave) = Concepto_Utilidad Then
                                        dImporte = Calculo_dImporte(Array_I)
'                                        dImporte = 0: dImporte = Round((Val(Array_I(2, rs_MGenera.AbsolutePosition)) * Array_CC(2, i)) / 100, 2)
'                                        If UBound(Array_CC) = i Then dImporte = Val(Array_I(3, rs_MGenera.AbsolutePosition)) Else Array_I(3, rs_MGenera.AbsolutePosition) = Val(Array_I(3, rs_MGenera.AbsolutePosition)) - dImporte
                                        
                                        
                                        Call Graba_Asiento_Contable(wcia, Trim(rs_MGenera_Boletas!PlaCod), TxtAno, s_MesSeleccion, _
                                        Trim(rs_MGenera_Boletas!semana), i_Numero_VoucherG, Cuenta_dUtilidades, "1", "", 0, dImporte, s_Tipo_Boleta, _
                                        s_TipoTrabajador_Report, mCenCosto, MArea)
                                        Call Graba_Asiento_Contable(wcia, Trim(rs_MGenera_Boletas!PlaCod), TxtAno, s_MesSeleccion, _
                                        Trim(rs_MGenera_Boletas!semana), i_Numero_VoucherG, Cuenta_dUtilidades, "0", "", dImporte, 0, s_Tipo_Boleta, _
                                        s_TipoTrabajador_Report, mCenCosto, MArea)
                                    Else
                                        dImporte = Calculo_dImporte(Array_I)
                                        Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                                        rs_MGenera_Boletas!semana, i_Numero_VoucherG, rs_MGenera2!plancta_Cargo1, "", dImporte, 1, "0", mCenCosto, MArea)
                                    End If
                                End If
                                Set rs_MGenera2 = Nothing
                                Call Cerrar_Conexion_Base_Datos_Access
                            Case 2
                                dImporte = Calculo_dImporte(Array_I)
                                If dImporte < 0 Then
                                    Suma_Cuenta_79_negativo = Suma_Cuenta_79_negativo + Abs(dImporte)
                                Else
                                    Suma_Cuenta_79 = Suma_Cuenta_79 + dImporte
                                End If
                                If UCase(Trim(rs_MGenera!Clave)) <> UCase(Trim(Concepto_Utilidad)) Then
                                    
                                    If Len(rs_MGenera!cma_ctacontable) > 4 Then
                                        '------------------
                                        'RODA: GTA - 26/04/2012
                                        'EN LASO LA CUENTA SE MAYO DE 4 DIGITOS GENERA ASIENTO DE DEBE Y HABER POR MONTO ASINGADO
                                        '-----------------------------------------------------
                                        If Len(rs_MGenera!cma_ctacontable) > 5 Then
                                            'Debe
                                            Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                                            rs_MGenera_Boletas!semana, i_Numero_VoucherG, rs_MGenera!cma_ctacontable, "", dImporte, 1, "1", mCenCosto, MArea)
                                                                                    
                                            'Haber
                                            Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                                            rs_MGenera_Boletas!semana, i_Numero_VoucherG, rs_MGenera!cma_ctacontable, "", dImporte, 2, "0", mCenCosto, MArea)
                                            '-----------------------------------------------------
                                        Else
                                           Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                                            rs_MGenera_Boletas!semana, i_Numero_VoucherG, s_CentroCosto & rs_MGenera!cma_ctacontable, "", dImporte, 1, "1", mCenCosto, MArea)
                                        End If
                                    Else
                                        Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                                        rs_MGenera_Boletas!semana, i_Numero_VoucherG, s_CentroCosto & rs_MGenera!cma_ctacontable, "", dImporte, 1, "1", mCenCosto, MArea)
                                    End If
                                    
                                Else
'                                    Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PLACOD, txtano, s_MesSeleccion, _
'                                    rs_MGenera_Boletas!semana, i_Numero_VoucherG, Cuenta_dUtilidades, "", rs_MGenera!total_montos, 1, "1", mCenCosto)
'                                    Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PLACOD, txtano, s_MesSeleccion, _
'                                    rs_MGenera_Boletas!semana, i_Numero_VoucherG, Cuenta_dUtilidades, "", rs_MGenera!total_montos, 1, "0", mCenCosto)
                                    Call Graba_Asiento_Contable(wcia, Trim(rs_MGenera_Boletas!PlaCod), TxtAno, s_MesSeleccion, _
                                    Trim(rs_MGenera_Boletas!semana), i_Numero_VoucherG, Cuenta_dUtilidades, "1", "", 0, dImporte, s_Tipo_Boleta, _
                                    s_TipoTrabajador_Report, mCenCosto, MArea)
                                    Call Graba_Asiento_Contable(wcia, Trim(rs_MGenera_Boletas!PlaCod), TxtAno, s_MesSeleccion, _
                                    Trim(rs_MGenera_Boletas!semana), i_Numero_VoucherG, Cuenta_dUtilidades, "0", "", dImporte, 0, s_Tipo_Boleta, _
                                    s_TipoTrabajador_Report, mCenCosto, MArea)
                                End If
                        End Select
                        rs_MGenera.MoveNext
                End Select
            Loop

        End If
        Set rs_MGenera = Nothing
    Next i_Contador
End Sub

Sub Proceso_Descuentos_Asiento_Naturaleza_Obreros()
    Dim i_Contador As Integer
    For i_Contador = 1 To 2
        Call Recupera_Codigos_Conceptos_Ingresos("D", s_TipoTrabajador, wcia)
        Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
        If bBoleano = False Then Call Arma_Arreglo(Array_D, rs_MGenera)
        Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
        If rs_MGenera.EOF = False Then
            rs_MGenera.MoveFirst
            Do While Not rs_MGenera.EOF
                Debug.Print rs_MGenera!Clave
                Select Case i_Contador
                    Case 1
                    Call Graba_Informacion_Total(Trim(rs_MGenera!Clave), rs_MGenera_Boletas!FechaProceso, Trim(s_Tipo_Boleta), Trim(s_TipoTrabajador), Trim(rs_MGenera_Boletas!PlaCod), Trim(rs_MGenera_Boletas!semana), wcia)
'                        Call Graba_Informacion_Totales(rs_MGenera!Clave, txtano, s_MesSeleccion, _
'                        s_Tipo_Boleta, s_TipoTrabajador, rs_MGenera_Boletas!PLACOD, rs_MGenera_Boletas!semana, wcia)
                    Case 2
                        dImporte = Calculo_dImporte(Array_D)
                        If s_CtaIpss = rs_MGenera!cma_ctacontable Then
                            Call Ejecuta_Asientos_Contables2(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, rs_MGenera_Boletas!semana, _
                            i_Numero_VoucherG, rs_MGenera!cma_ctacontable, "", i_Monto_Ipss, 2, "0", mCenCosto, MArea)
                        Else
                            Call Ejecuta_Asientos_Contables2(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, rs_MGenera_Boletas!semana, _
                            i_Numero_VoucherG, rs_MGenera!cma_ctacontable, "", dImporte, 2, "0", mCenCosto, MArea)
                        End If
                End Select
                rs_MGenera.MoveNext
            Loop
        End If
        Set rs_MGenera = Nothing
    Next i_Contador
End Sub

Sub Proceso_Aportes_Asiento_Naturaleza_Obreros(Opcion As Integer)
    Dim i_Contador As Integer
    For i_Contador = 1 To 2
        Call Recupera_Codigos_Conceptos_Ingresos("A", s_TipoTrabajador, wcia)
        Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
        If bBoleano = False Then Call Arma_Arreglo(Array_A, rs_MGenera)
        Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
        If rs_MGenera.EOF = False Then
            rs_MGenera.MoveFirst
            Do While Not rs_MGenera.EOF
                Debug.Print rs_MGenera!Clave
                Select Case i_Contador
                    Case 1
                        Call Graba_Informacion_Total(Trim(rs_MGenera!Clave), rs_MGenera_Boletas!FechaProceso, Trim(s_Tipo_Boleta), Trim(s_TipoTrabajador), Trim(rs_MGenera_Boletas!PlaCod), Trim(rs_MGenera_Boletas!semana), wcia)
'                        Call Graba_Informacion_Totales(rs_MGenera!Clave, txtano, s_MesSeleccion, _
'                        s_Tipo_Boleta, s_TipoTrabajador, rs_MGenera_Boletas!PLACOD, rs_MGenera_Boletas!semana, wcia)
                    Case 2
                        Select Case Opcion
                            Case 1
                                Call Conectar_Base_Datos_Access(s_CodEmpresa_Starsoft)
                                Call Recuperar_Cuentas_Naturaleza_Transferencia(s_CentroCosto & rs_MGenera!cma_ctacontable)
                                Set rs_MGenera2 = CompRetenciones.rs_compRetenciones_Pub
                                Set CompRetenciones.rs_compRetenciones_Pub = Nothing
                                'AKI AGTREO ESTA FILA DEL MONTO IPS ANTES DEL IF PARA K EN EL SIGUIENTE PROC GRABE LA 40
                                dImporte = Calculo_dImporte(Array_A)
                                i_Monto_Ipss = dImporte
                                If rs_MGenera2.EOF = False Then
                                    i_Monto_Ipss = dImporte
                                    Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                                    rs_MGenera_Boletas!semana, i_Numero_VoucherG, rs_MGenera2!plancta_Cargo1, "", dImporte, 1, "0", mCenCosto, MArea)
                                End If
                                Set rs_MGenera2 = Nothing
                                Call Cerrar_Conexion_Base_Datos_Access
                            Case 2
                                dImporte = Calculo_dImporte(Array_A)
                                If dImporte < 0 Then
                                    Suma_Cuenta_79_negativo = Suma_Cuenta_79_negativo + Abs(dImporte)
                                Else
                                    Suma_Cuenta_79 = Suma_Cuenta_79 + dImporte
                                End If
                                Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, rs_MGenera_Boletas!semana, _
                                i_Numero_VoucherG, s_CentroCosto & rs_MGenera!cma_ctacontable, "", dImporte, 1, "1", mCenCosto, MArea)
                        End Select
                End Select
                rs_MGenera.MoveNext
            Loop
        End If
        Set rs_MGenera = Nothing
    Next i_Contador
End Sub

Sub Envia_Asiento_Neto()
    Dim r_Monto_Neto As Double
    Call Recupera_Monto_Neto(TxtAno, s_MesSeleccion, s_Tipo_Boleta, rs_MGenera_Boletas!PlaCod, rs_MGenera_Boletas!semana, rs_MGenera_Boletas!FechaProceso)
    Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
    If bBoleano = False Then
        If Not rs_MGenera.EOF Then
            Erase Array_N
            ReDim Array_N(1, 2)
            Array_N(1, 1) = rs_MGenera!totneto
            Array_N(1, 2) = rs_MGenera!totneto
        End If
    End If
    dImporte = 0: dImporte = Round((Val(Array_N(1, 1)) * Array_CC(I)) / 100, 2)
    If UBound(Array_CC) = I Then
        dImporte = Val(Array_N(1, 2))
    Else
        Array_N(1, 2) = Val(Array_N(1, 2)) - dImporte
    End If
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    r_Monto_Neto = dImporte 'rs_MGenera!TOTNETO
    Set rs_MGenera = Nothing
    If Verifica_Existencia_Registro(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, _
        s_MesSeleccion, rs_MGenera_Boletas!semana, i_Numero_VoucherG, s_Asiento_Neto, s_Tipo_Boleta, mCenCosto, "0") = False Then
        Call Graba_Asiento_Contable(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
        rs_MGenera_Boletas!semana, i_Numero_VoucherG, s_Asiento_Neto, 0, "", dImporte, _
        0, s_Tipo_Boleta, s_TipoTrabajador_Report, mCenCosto, MArea)
    Else
    '    Call Graba_Asiento_Contable(wcia, rs_MGenera_Boletas!PLACOD, txtano, s_MesSeleccion, _
     '   rs_MGenera_Boletas!semana, i_Numero_VoucherG, s_Asiento_Neto, 0, "", r_Monto_Neto, _
      '  0, s_Tipo_Boleta, s_TipoTrabajador_Report)
        '*********modificar para que la edicion vaya al haber*************
        Call Editar_Asiento_Contable2(dImporte, wcia, rs_MGenera_Boletas!PlaCod, _
        TxtAno, s_MesSeleccion, rs_MGenera_Boletas!semana, i_Numero_VoucherG, s_Asiento_Neto, s_Tipo_Boleta, mCenCosto, MArea)
        '*****************************************************************
    End If
End Sub

Sub Recupera_CtaContable_Asiento_Neto()
    Select Case CmbBoleta.Text
        Case "NORMAL": s_Envio_Clave_Boleta = "001"
        Case "GRATIFICACION": s_Envio_Clave_Boleta = "003"
        Case "VACACIONES": s_Envio_Clave_Boleta = "002"
        Case "UTILIDADES": s_Envio_Clave_Boleta = "011"
    End Select
    Call Recupera_Asiento_Contable_Neto(s_Envio_Clave_Boleta, s_TipoTrabajador)
    s_Asiento_Neto = Crear_Plan_Contable.s_CtaContableNeto
End Sub

Sub Ejecuta_Asientos_Contables(CodCompañia As String, CodTrabajador As String, Año_Proceso As _
String, MesSeleccion As String, SemProceso As String, Voucher As String, CtaContable As String, _
DescCta As String, MontoInt As Double, Opcion As Integer, TipoAsiento As String, _
CentroCosto As String, Area As String)
    If Verifica_Existencia_Registro(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
    SemProceso, Voucher, CtaContable, s_Tipo_Boleta, CentroCosto, TipoAsiento) = False Then
        Select Case Opcion
            Case 1
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, 0, MontoInt, s_Tipo_Boleta, _
                s_TipoTrabajador_Report, CentroCosto, Area)
            Case 2
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, MontoInt, 0, s_Tipo_Boleta, _
                  s_TipoTrabajador_Report, CentroCosto, Area)
        End Select
    Else
        Call Editar_Asiento_Contable(MontoInt, CodCompañia, CodTrabajador, Año_Proceso, _
        MesSeleccion, SemProceso, Voucher, CtaContable, s_Tipo_Boleta, CentroCosto, Area)
    End If
End Sub


'*****25/06/2009
'*****JJ
'*****Sumar las deducciones cta cta + otros descuentos
'*****
Sub Ejecuta_Asientos_Contables2(CodCompañia As String, CodTrabajador As String, Año_Proceso As _
String, MesSeleccion As String, SemProceso As String, Voucher As String, CtaContable As String, _
DescCta As String, MontoInt As Double, Opcion As Integer, TipoAsiento As String, _
CentroCosto As String, Area As String)
    If Verifica_Existencia_Registro(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
    SemProceso, Voucher, CtaContable, s_Tipo_Boleta, CentroCosto, TipoAsiento) = False Then
        If CtaContable = "1419102" And Left(CodTrabajador, 1) = "E" Then CtaContable = "1419101"
        If CtaContable = "1419101" And Left(CodTrabajador, 1) = "O" Then CtaContable = "1419102"
        Select Case Opcion
            Case 1
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, 0, MontoInt, s_Tipo_Boleta, _
                s_TipoTrabajador_Report, CentroCosto, Area)
            Case 2
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, MontoInt, 0, s_Tipo_Boleta, _
                s_TipoTrabajador_Report, CentroCosto, Area)
        End Select
    Else
        Call Editar_Asiento_Contable2(MontoInt, CodCompañia, CodTrabajador, Año_Proceso, _
        MesSeleccion, SemProceso, Voucher, CtaContable, s_Tipo_Boleta, CentroCosto, Area)
    End If
End Sub

Sub Enviar_Cuenta_79()
    Call Recupera_Codigos_Conceptos_Ingresos("A", s_TipoTrabajador, wcia)
    Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    If rs_MGenera.EOF = False Then
        rs_MGenera.MoveFirst
        Call Conectar_Base_Datos_Access(s_CodEmpresa_Starsoft)
        Call Recuperar_Cuentas_Naturaleza_Transferencia(s_CentroCosto & rs_MGenera!cma_ctacontable)
        Set rs_MGenera2 = CompRetenciones.rs_compRetenciones_Pub
        Set CompRetenciones.rs_compRetenciones_Pub = Nothing
        If rs_MGenera2.EOF = False Then
            Call Ejecuta_Asientos_Contables(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
            rs_MGenera_Boletas!semana, i_Numero_VoucherG, rs_MGenera2!plancta_abono1, "", Suma_Cuenta_79, 2, "1", mCenCosto, MArea)
            If Suma_Cuenta_79_negativo > 0 Then
                Call Graba_Asiento_Contable(wcia, rs_MGenera_Boletas!PlaCod, TxtAno, s_MesSeleccion, _
                rs_MGenera_Boletas!semana, i_Numero_VoucherG, rs_MGenera2!plancta_abono1, "1", "", 0, Suma_Cuenta_79_negativo, s_Tipo_Boleta, _
                s_TipoTrabajador_Report, mCenCosto, MArea)
            End If
        End If
        Set rs_MGenera2 = Nothing
        Call Cerrar_Conexion_Base_Datos_Access
    End If
    Set rs_MGenera = Nothing
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    'Call Proceso_Integral
    If wGrupoPla = "01" Then P1.Visible = False
End Sub

Sub Recupera_Cuenta_Asociada()
    '********por falta de tiempo se utilizo este codigo, pero hay que mejorarlo para*********
    '***************que no tenga que realizarse el bucle del for*****************************
    Dim i_Contador As Integer
    Call Recupera_Informacion_Contable_Fijo
    Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    rs_MGenera.MoveFirst
    For i_Contador = 1 To 5
        Select Case rs_MGenera!cmaf_clave
            Case "004": s_CtaIpss = rs_MGenera!cmaf_ctacontable
        End Select
        rs_MGenera.MoveNext
    Next i_Contador
    Set rs_MGenera = Nothing
End Sub

Sub Recupera_Cuenta_Asociada_2()
    '********por falta de tiempo se utilizo este codigo, pero hay que mejorarlo para*********
    '***************que no tenga que realizarse el bucle del for*****************************
    Dim i_Contador As Integer
    Call Recupera_Informacion_Contable_Fijo
    Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    rs_MGenera.MoveFirst
    For i_Contador = 1 To 5
        Select Case rs_MGenera!cmaf_clave
            Case "005": s_CtaIpss = rs_MGenera!cmaf_ctacontable
        End Select
        rs_MGenera.MoveNext
    Next i_Contador
    Set rs_MGenera = Nothing
End Sub

Sub Trae_Cuenta_Asociada(ByVal tipo As String)
    Dim rsTemp As ADODB.Recordset
    Call Recupera_Informacion_Contable_Fijo("")
    Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    If Not rs_MGenera.EOF Then
        rs_MGenera.MoveFirst
        Do While Not rs_MGenera.EOF
            Select Case rs_MGenera!cmaf_clave
                Case "004": If tipo = "02" Then s_CtaIpss = rs_MGenera!cmaf_ctacontable: Exit Do
                Case "005": If tipo = "01" Then s_CtaIpss = rs_MGenera!cmaf_ctacontable: Exit Do
            End Select
            rs_MGenera.MoveNext
        Loop
    End If
    Set rs_MGenera = Nothing
End Sub

Sub Genera_Numero_Voucher()
    i_Numero_Voucher = i_Numero_Voucher + 1
    Select Case i_Numero_Voucher
        Case Is < 10: i_Numero_VoucherG = "000" & i_Numero_Voucher
        Case Is < 100: i_Numero_VoucherG = "00" & i_Numero_Voucher
        Case Is < 1000: i_Numero_VoucherG = "0" & i_Numero_Voucher
        Case Is < 10000: i_Numero_VoucherG = i_Numero_Voucher
    End Select
End Sub

Sub Selecciona_Tipo_Boleta()
    Select Case CmbBoleta.Text
        Case "NORMAL": s_Tipo_Boleta = "01"
        Case "GRATIFICACION": s_Tipo_Boleta = "03"
        Case "VACACIONES": s_Tipo_Boleta = "02"
        Case "LIQUIDACION": s_Tipo_Boleta = "04"
        Case "UTILIDADES": s_Tipo_Boleta = "11"
        Case "EXTORNO VACACIONES": s_Tipo_Boleta = "02"
'        Case "CTS": s_Tipo_Boleta = "04"
'        '*********CODIGO AGREGADO GIOVANNI 20082007****************
'        Case "PROV CTS": s_Tipo_Boleta = "14"
'        Case "PROV GRATIFICACION": s_Tipo_Boleta = "13"
'        Case "PROV VACACIONES": s_Tipo_Boleta = "12"
        '**********************************************************
    End Select
End Sub

Sub Recupera_tipo_trabajador()
    s_TipoTrabajador = Empty
    s_TipoTrabajador = Trim(fc_CodigoComboBox(CmbTrabTipo, 2))
    Select Case s_TipoTrabajador
        Case "02": s_TipoTrabajador_Report = 0
        Case "01": s_TipoTrabajador_Report = 1
    End Select
End Sub

Sub Proceso_Central()
    Dim i_Contador As Integer
    i_Numero_Voucher = 0: i_Contador = 0
    
    If MsgBox("Desea generar cuentas destino?", vbYesNo + vbQuestion, "Sistema") = vbYes Then
        Bol_ConDestino = True
    Else
        Bol_ConDestino = False
    End If
    
    
    'MODIFICACION REALIZADA PARA TRABAJAR CON n CENTRO DE COSTO

    Call Recupera_Informacion_Total_Boletas(TxtAno.Text, s_MesSeleccion, s_Tipo_Boleta, wcia, s_TipoTrabajador)
    Set rs_MGenera_Boletas = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    If rs_MGenera_Boletas.EOF = False Then
        Call Codigo_Empresa_Starsoft
        If Trim(s_CodEmpresa_Starsoft) = "0" Then
            MsgBox "No se encuentra el código de empresa externo."
            Exit Sub
        End If
        If rs_MGenera_Boletas.RecordCount > 1 Then
            P1.Min = 1: P1.Max = rs_MGenera_Boletas.RecordCount
        End If
        rs_MGenera_Boletas.MoveFirst
        Do While Not rs_MGenera_Boletas.EOF
            DoEvents
            Call Genera_Numero_Voucher
            
            Call Crea_Array_CC
            bBoleano = False
            For I = LBound(Array_CC) To UBound(Array_CC)
                Centro_Costo = Empty: Por_CC = 0: dImporte = 0
                Centro_Costo = Trim(rs_MGenera_Boletas.Fields.Item("CCOSTO" & I))
                Por_CC = rs_MGenera_Boletas.Fields.Item("PORC" & I)
                If Trim(Centro_Costo) = "" Then Exit For
                
                If Recupera_Centro_Costo(wcia, Centro_Costo) = True Then
                    s_CentroCosto = Trim(Crear_Plan_Contable.s_CentroCostoPub)
                    mCenCosto = Trim(Crear_Plan_Contable.mCentroCosto)
                    'aca mgirao cargar area
                    MArea = rs_MGenera_Boletas!Area
                 
                    Suma_Cuenta_79 = 0
                    Suma_Cuenta_79_negativo = 0
                    Call Recupera_CtaContable_Asiento_Neto
                    Call Trae_Cuenta_Asociada(s_TipoTrabajador)
                    Call Proceso_Ingresos_Asiento_Naturaleza_Obreros(1)
                    Call Proceso_Aportes_Asiento_Naturaleza_Obreros(1)
                    Call Proceso_Descuentos_Asiento_Naturaleza_Obreros
                    Call Envia_Asiento_Neto
                    'Moficacion no genera naturaleza y  trans
                    Call Proceso_Ingresos_Asiento_Naturaleza_Obreros(2)
                    Call Proceso_Aportes_Asiento_Naturaleza_Obreros(2)
                    Call Enviar_Cuenta_79
                Else
                    Call Llena_Barra
                    MsgBox "Falta el Centro de Costo del Personal: " & rs_MGenera_Boletas!PlaCod
                    Exit Do
                End If
                bBoleano = True
            Next
            If rs_MGenera_Boletas.RecordCount > 1 Then
                i_Contador = i_Contador + 1
                On Error Resume Next
                P1.Value = i_Contador
            End If
            rs_MGenera_Boletas.MoveNext
        Loop
    End If
    MsgBox "Proceso Terminado", vbInformation
End Sub
Sub Proceso_General()

    If Val(TxtAno.Text) < 1 Then
        MsgBox "Ingrese año de proceso", vbExclamation
        TxtAno.SetFocus
        Exit Sub
    End If
    '@01 JCJS 07022023 Se modificó la logica a <0, ya que el primer valor siempre es 0
        If Me.CmbMes.ListIndex < 0 Then
        MsgBox "Seleccione mes del proceso", vbExclamation
        CmbMes.SetFocus
        Exit Sub
    End If
    
    If TxtAno.Text < 2013 Then
       MsgBox "En este sistema se pueden trabajar desde Setiembre del 2013", vbExclamation
       TxtAno.SetFocus
       Exit Sub
    End If
    
    If TxtAno.Text = 2013 And CmbMes.ListIndex + 1 < 9 Then
       MsgBox "En este sistema se pueden trabajar desde Setiembre del 2013", vbExclamation
       TxtAno.SetFocus
       Exit Sub
    End If
    
    If CmbTrabTipo.ListIndex < 1 Then
        MsgBox "Seleccione tipo de trabajador", vbExclamation
        CmbTrabTipo.SetFocus
        Exit Sub
    End If

    Selecciona_Tipo_Boleta
    Dim lExtorno As String
    lExtorno = "N"
    If CmbBoleta.Text = "EXTORNO VACACIONES" Then lExtorno = "S"
    Sql = "uSp_Pla_Asiento '" & wcia & "'," & TxtAno.Text & "," & CmbMes.ListIndex + 1 & ",'" & s_Tipo_Boleta & "','" & s_TipoTrabajador & "','" & wuser & "','" & wNamePC & "','" & lExtorno & "'"
    Debug.Print Sql
    If (fAbrRst(rs, Sql)) Then
       If Trim(rs(0) & "") = "OK" Then
          MsgBox "Asiento Generado Correctamente", vbInformation
          Screen.MousePointer = vbDefault
       ElseIf Left(Trim(rs(0) & ""), 2) = "NG" Then '131222
          MsgBox "Error al generar asiento" & Chr(13) & Trim(rs(0) & ""), vbCritical, Me.Caption
          Screen.MousePointer = vbDefault
       ElseIf Trim(rs!cod_mensaje) = "NO" Then
       
        Dim mensaje As String
        Dim codigoTrabajadorConca As String
        mensaje = rs!mensaje
       
       'obtener la lista de trabajadores observados
       rs.MoveFirst
        Do While Not rs.EOF
            If Len(Trim(codigoTrabajadorConca)) < 1 Then
                codigoTrabajadorConca = "'" + Trim(rs!cgaux) & "'"
            Else
                    codigoTrabajadorConca = codigoTrabajadorConca & ", '" & Trim(rs!cgaux) & "'"
            End If
            rs.MoveNext
        Loop
        
        'listar los datos del trabajador
        Sql = "select "
        Sql = Sql & "placod as tra_cod,"
        Sql = Sql & "(select top 1 rtrim(ltrim(ap_pat)) + ' ' + rtrim(ltrim(ap_mat)) + ', ' + rtrim(ltrim(nom_1)) + ' ' + rtrim(ltrim(nom_2)) from Planillas where status <>'*' and placod = ph.placod  ) as tra_ape_nom,"
        Sql = Sql & "(select descripcion  from Pla_CCostos where codigo=ph.ccosto1) as cen_costo_nom "
        Sql = Sql & " From "
        Sql = Sql & " plahistorico As ph "
        Sql = Sql & " Where "
        Sql = Sql & " status<>'*' "
        Sql = Sql & " and year(fechaproceso)=" & TxtAno.Text
        Sql = Sql & " and month(fechaproceso)=" & CmbMes.ListIndex + 1
        Sql = Sql & " and proceso='" & s_Tipo_Boleta & "' "
        Sql = Sql & " and tipotrab='" & s_TipoTrabajador & "' "
        Sql = Sql & " and placod in (" & codigoTrabajadorConca & ") "
       
       Dim lineaMensaje As String
       Dim contador As Integer
        If (fAbrRst(rs, Sql)) = True Then
            rs.MoveFirst
            Do While Not rs.EOF
                If Len(Trim(lineaMensaje)) < 1 Then
                    contador = 1
                    lineaMensaje = contador & ".- Codigo Trabajador: " + Trim(rs!tra_cod) + "-" + Trim(rs!tra_ape_nom) + " - Centro Costo: " + Trim(rs!cen_costo_nom)
                Else
                    lineaMensaje = lineaMensaje & Chr(13) & contador & ".- Codigo Trabajador: " + Trim(rs!tra_cod) + "-" + Trim(rs!tra_ape_nom) + " - Centro Costo: " + Trim(rs!cen_costo_nom)
                End If
                rs.MoveNext
                contador = contador + 1
            Loop
        End If
       
       mensaje = mensaje & Chr(13) & Chr(13) & lineaMensaje
          MsgBox mensaje, vbExclamation
          Screen.MousePointer = vbDefault
          rs.Close: Set rs = Nothing: Exit Sub
       Else
          MsgBox "Asiento No se Genero, Revise", vbExclamation
          rs.Close: Set rs = Nothing: Exit Sub
          Exit Sub
       End If
    End If
     
    Dim mLote As String
    Dim mVoucher As String
    Dim mTit As String

If s_TipoTrabajador = "01" Then
   If s_Tipo_Boleta = "01" Then
      mLote = "PEN": mVoucher = "0501000001"
      mTit = "PLANILLAS EMPLEADOS NORMAL"
   End If
   If s_Tipo_Boleta = "02" And lExtorno <> "S" Then
      mLote = "PEV": mVoucher = "0501000003"
      mTit = "PLANILLAS EMPLEADOS VACACIONES"
   End If
   If s_Tipo_Boleta = "03" Then
      mLote = "PEG": mVoucher = "0501000007"
      mTit = "PLANILLAS EMPLEADOS GRATIFICACION"
   End If
   If s_Tipo_Boleta = "02" And lExtorno = "S" Then
      mLote = "EVE": mVoucher = "0501000005"
      mTit = "EXTORNO EMPLEADOS VACACIONES"
   End If
   If s_Tipo_Boleta = "11" Then
      mLote = "PEU": mVoucher = "0501000009"
      mTit = "EMPLEADOS UTILIDADES"
   End If
End If
If s_TipoTrabajador = "02" Then
   If s_Tipo_Boleta = "01" Then
      mLote = "PON": mVoucher = "0501000002"
      mTit = "PLANILLAS OBREROS NORMAL"
   End If
   If s_Tipo_Boleta = "02" Then
      mLote = "POV": mVoucher = "0501000004"
      mTit = "PLANILLAS OBREROS VACACIONES"
   End If
   If s_Tipo_Boleta = "03" Then
      mLote = "POG": mVoucher = "0501000008"
      mTit = "PLANILLAS OBREROS GRATIFICACION"
   End If
   If s_Tipo_Boleta = "02" And lExtorno = "S" Then
      mLote = "EVO": mVoucher = "0501000006"
      mTit = "EXTORNO OBRERO VACACIONES"
   End If
   If s_Tipo_Boleta = "11" Then
      mLote = "POU": mVoucher = "0501000010"
      mTit = "OBREROS UTILIDADES"
   End If
End If
Call Carga_Asiento_Excel(TxtAno.Text, CmbMes.ListIndex + 1, mLote, mVoucher, mTit, CmbMes.Text, "S")
  
End Sub
Sub Proceso_Reporte_Detallado()
    If CmbTrabTipo.ListIndex = 0 Then Exit Sub
    s_Año_ProcesoReport = TxtAno: Call Captura_Mes_Seleccionado
    s_Mes_ProcesoReport = s_MesSeleccion ': Call Recupera_tipo_trabajador
    s_Tip_TrabajadorReport = s_TipoTrabajador_Report: Call Selecciona_Tipo_Boleta
    s_Tip_BoletaReport = s_Tipo_Boleta
    Select Case FrmMGenera.Caption
        'Case "Imprimir Asientos - DETALLADO": RptGeneralesCont.Show
        'Case Else: RptDetalles.Show
    End Select
End Sub

Sub Llena_Barra()
    Dim i_Contador As Integer
    P1.Min = 1: P1.Max = 10
    For i_Contador = 1 To 10: P1.Value = i_Contador: Next i_Contador
End Sub

Sub Codigo_Empresa_Starsoft()
    'Call Recuperar_Codigo_Empresa_Starsoft(wcia)
    Call Trae_Cia_StarSoft(wcia, TxtAno.Text)
    Set rs_MGenera3 = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    s_CodEmpresa_Starsoft = rs_MGenera3!EMP_ID
    Set rs_MGenera3 = Nothing
End Sub


