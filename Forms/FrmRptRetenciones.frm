VERSION 5.00
Begin VB.Form FrmRptRetenciones 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobante de Retenciones "
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4200
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1065
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3990
      Begin VB.TextBox TxtCodTrabajador 
         Height          =   315
         Left            =   2250
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox TxtAño 
         Height          =   285
         Left            =   2250
         TabIndex        =   1
         Top             =   225
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese Año Proceso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   1560
      End
   End
End
Attribute VB_Name = "FrmRptRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_Comp_Retenciones As ADODB.Recordset
Dim rs_Comp_Retenciones2 As ADODB.Recordset
Dim rs_Comp_Retenciones3 As ADODB.Recordset
Public s_Jornal As String
Public s_Incremento As String
Public s_Gratificacion As String
Public s_Vacaciones As String
Public i_Total_General1 As Double
Public s_NombreCompleto As String
Public i_TotalRetenciones As Double
Public s_NombreAfp As String
Dim i_TotRemunAseg As Double
Dim r_FactorPrimaseguro_adi As Double
Dim r_FactorComisionRA_adi As Double
Dim i_TotalComisionRA_adi As Double
Dim i_TotalPrimaSeguro_adi As Double
Dim i_Factor_TotalSNP_adi As Double
Dim i_TotAportObligatorio_adi As Double
Dim r_FactorAporObligatorio_adi As Double
Dim r_TotalAporSnp As Double
Public i_TotRetRetribucion_Onp As Double
Public i_TotRetRetribucion_Afp As Double
Dim r_CantidadTope_adi As Double
Dim s_MesSeleccion As String
Sub Recupera_Jornal()
    Call Recupera_Jornal_Basico(wcia, TxtAño, TxtCodTrabajador, "01")
    Set rs_Comp_Retenciones = CompRetenciones.rs_compRetenciones_Pub
    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
    If IsNull(rs_Comp_Retenciones!Total_BolNormal) Then
        s_Jornal = "0"
    Else
        s_Jornal = rs_Comp_Retenciones!Total_BolNormal
    End If
    Set rs_Comp_Retenciones = Nothing
End Sub
Sub Recupera_Incremento_Monto()
    Call Recupera_IncrementoAfp(wcia, TxtAño, TxtCodTrabajador)
    Set rs_Comp_Retenciones = CompRetenciones.rs_compRetenciones_Pub
    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
    If IsNull(rs_Comp_Retenciones!Total_IncrementoAFP) Then
        s_Incremento = "0"
    Else
        s_Incremento = rs_Comp_Retenciones!Total_IncrementoAFP
    End If
    Set rs_Comp_Retenciones = Nothing
End Sub
Sub Recupera_Gratificacion_Monto()
    Call Recupera_Gratificaciones(wcia, TxtAño, TxtCodTrabajador, "03")
    Set rs_Comp_Retenciones = CompRetenciones.rs_compRetenciones_Pub
    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
    If IsNull(rs_Comp_Retenciones!Total_BolGratificacion) Then
        s_Gratificacion = "0"
    Else
        s_Gratificacion = rs_Comp_Retenciones!Total_BolGratificacion
    End If
    Set rs_Comp_Retenciones = Nothing
End Sub
Sub Recupera_Vacaciones_Monto()
    Call Recupera_Vacaciones(wcia, TxtAño, TxtCodTrabajador, "02")
    Set rs_Comp_Retenciones = CompRetenciones.rs_compRetenciones_Pub
    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
    If IsNull(rs_Comp_Retenciones!Total_BolVacaciones) Then
        s_Vacaciones = "0"
    Else
        s_Vacaciones = rs_Comp_Retenciones!Total_BolVacaciones
    End If
    Set rs_Comp_Retenciones = Nothing
End Sub
Sub Nombre_Trabajador()
    Call Recupera_Nombre_Trabajador(wcia, TxtCodTrabajador)
    Set rs_Comp_Retenciones = CompRetenciones.rs_compRetenciones_Pub
    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
    s_NombreCompleto = rs_Comp_Retenciones!nombres
    Set rs_Comp_Retenciones = Nothing
End Sub
Sub Proceso_Calculo_Montos_Finales()
    Dim s_CodTrabajador_Monto As String
    Dim s_Mes_Proceso_Monto As String
    s_CodTrabajador_Monto = "": s_Mes_Proceso_Monto = ""
    Call Recupera_Informacion_Total(wcia, TxtAño, TxtCodTrabajador)
    Set rs_Comp_Retenciones = CompRetenciones.rs_compRetenciones_Pub
    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
    If rs_Comp_Retenciones.EOF = False Then
        rs_Comp_Retenciones.MoveFirst
        Do While rs_Comp_Retenciones.EOF = False
            If s_CodTrabajador_Monto <> rs_Comp_Retenciones!PlaCod And s_Mes_Proceso_Monto <> Month(rs_Comp_Retenciones!FechaProceso) Then
                Call Recupera_Mes_Seleccion(Month(rs_Comp_Retenciones!FechaProceso))
                Call Monto_Total_Remuneracion_Asegurable(wcia, Year(rs_Comp_Retenciones!FechaProceso), _
                s_MesSeleccion, rs_Comp_Retenciones!CodAfp)
                Call Recupera_Factores(wcia, Year(rs_Comp_Retenciones!FechaProceso), _
                s_MesSeleccion, rs_Comp_Retenciones!CodAfp)
                Call Calculamos_Montos_Totales(rs_Comp_Retenciones!CodAfp)
            End If
            s_CodTrabajador_Monto = rs_Comp_Retenciones!PlaCod
            s_Mes_Proceso_Monto = Month(rs_Comp_Retenciones!FechaProceso)
            rs_Comp_Retenciones.MoveNext
        Loop
        i_TotalRetenciones = 0
        i_TotalRetenciones = Format(i_TotRetRetribucion_Onp + i_TotRetRetribucion_Afp, "0.00")
    End If
    rs_Comp_Retenciones.MoveLast
    If rs_Comp_Retenciones!CodAfp <> "01" And rs_Comp_Retenciones!CodAfp <> "02" Then
        Call Recupera_Nombre_Afp(rs_Comp_Retenciones!CodAfp)
        Set rs_Comp_Retenciones3 = Reportes_Centrales.rs_RptCentrales_pub
        Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
        s_NombreAfp = rs_Comp_Retenciones3!DESCRIP
        Set rs_Comp_Retenciones3 = Nothing
    Else
        s_NombreAfp = ""
    End If
    
    Set rs_Comp_Retenciones = Nothing
End Sub
Sub Monto_Total_Remuneracion_Asegurable(CodEmpresa As String, AñoProceso As String, MesProceso As _
String, CodPensiones As String)
    i_TotRemunAseg = 0
    Call Recuperar_Monto_Total_Remuneracion_Asegurable(CodEmpresa, AñoProceso, MesProceso, CodPensiones)
    Set rs_Comp_Retenciones2 = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_Comp_Retenciones2.EOF = False Then
        If IsNull(rs_Comp_Retenciones2!total_remaseg) Then
            i_TotRemunAseg = 0
        Else
            i_TotRemunAseg = rs_Comp_Retenciones2!total_remaseg
        End If
    Else
        i_TotRemunAseg = 0
    End If
    Set rs_Comp_Retenciones2 = Nothing
End Sub
Sub Recupera_Factores(CodEmpresa As String, AñoProceso As String, MesProceso As String, _
CodPensiones As String)
    Select Case CodPensiones
        Case "01": i_Factor_TotalSNP_adi = 13
        Case "02": i_Factor_TotalSNP_adi = 13
        Case Else
            Call Recupera_Factores_Operacionales(CodEmpresa, AñoProceso, MesProceso, CodPensiones)
            Set rs_Comp_Retenciones2 = Reportes_Centrales.rs_RptCentrales_pub
            Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
            If rs_Comp_Retenciones2.EOF = False Then
                r_FactorPrimaseguro_adi = rs_Comp_Retenciones2!Fac_PrimaSeg
                r_FactorComisionRA_adi = rs_Comp_Retenciones2!Fac_ComRA
                r_FactorAporObligatorio_adi = rs_Comp_Retenciones2!Fac_AporOblig
                r_CantidadTope_adi = rs_Comp_Retenciones2!tope
            Else
                r_FactorPrimaseguro_adi = 1: r_FactorComisionRA_adi = 1
                r_FactorAporObligatorio_adi = 1: r_CantidadTope_adi = 1
                Call Recupera_Nombre_Afp(CodPensiones)
                Set rs_Comp_Retenciones3 = Reportes_Centrales.rs_RptCentrales_pub
                Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
                MsgBox "Los Factores de Calculo para el Periodo " & AñoProceso & MesProceso & " " & _
                "para la AFP " & rs_Comp_Retenciones3!DESCRIP & " no ha sido ingresado ", vbCritical
                Set rs_Comp_Retenciones3 = Nothing
            End If
            Set rs_Comp_Retenciones2 = Nothing
    End Select
End Sub
Sub Calculamos_Montos_Totales(CodAfp As String)
    i_TotRetRetribucion_Onp = 0: i_TotRetRetribucion_Afp = 0
    Select Case CodAfp
        Case "01": r_TotalAporSnp = (i_TotRemunAseg * i_Factor_TotalSNP_adi) / 100
        Case "02": r_TotalAporSnp = (i_TotRemunAseg * i_Factor_TotalSNP_adi) / 100
        Case Else
            i_TotAportObligatorio_adi = (i_TotRemunAseg * r_FactorAporObligatorio_adi) / 100
            i_TotalComisionRA_adi = (i_TotRemunAseg * r_FactorComisionRA_adi) / 100
            Call Calcula_Prima_Seguro
    End Select
    Select Case CodAfp
        Case "01": i_TotRetRetribucion_Onp = Format(i_TotRetRetribucion_Onp + r_TotalAporSnp, "0.00")
        Case "02": i_TotRetRetribucion_Onp = Format(i_TotRetRetribucion_Onp + r_TotalAporSnp, "0.00")
        Case Else: i_TotRetRetribucion_Afp = Format(i_TotRetRetribucion_Afp + (i_TotalComisionRA_adi + i_TotalPrimaSeguro_adi), "0.00")
    End Select
End Sub
Sub Calcula_Prima_Seguro()
    If i_TotRemunAseg < r_CantidadTope_adi Then
        i_TotalPrimaSeguro_adi = (i_TotRemunAseg * r_FactorPrimaseguro_adi) / 100
    Else
        i_TotalPrimaSeguro_adi = (r_CantidadTope_adi * r_FactorPrimaseguro_adi) / 100
    End If
End Sub
Sub Recupera_Mes_Seleccion(MesProceso As Integer)
    Select Case MesProceso
        Case 1: s_MesSeleccion = "01": Case 2: s_MesSeleccion = "02"
        Case 3: s_MesSeleccion = "03": Case 4: s_MesSeleccion = "04"
        Case 5: s_MesSeleccion = "05": Case 6: s_MesSeleccion = "06"
        Case 7: s_MesSeleccion = "07": Case 8: s_MesSeleccion = "08"
        Case 9: s_MesSeleccion = "09": Case 10: s_MesSeleccion = "10"
        Case 11: s_MesSeleccion = "11": Case 12: s_MesSeleccion = "12"
    End Select
End Sub
Sub Proceso_Reporte_Central()
    Call Recupera_Jornal
    Call Recupera_Incremento_Monto
    Call Recupera_Gratificacion_Monto
    Call Recupera_Vacaciones_Monto
    Call Nombre_Trabajador
    i_Total_General1 = Val(s_Jornal) + Val(s_Incremento) + Val(s_Gratificacion) + Val(s_Vacaciones)
    Call Proceso_Calculo_Montos_Finales
    i_Direccion_Reportes = 7
    'RptAsistencias.Show
End Sub




