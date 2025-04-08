VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form FrmRptAportes 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes Aportes AFP"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4230
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1815
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   3990
      Begin VB.ComboBox Cmbafp 
         Height          =   315
         Left            =   2325
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   1515
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmRptAportes.frx":0000
         Left            =   2325
         List            =   "FrmRptAportes.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox TxtAño 
         Height          =   285
         Left            =   2325
         TabIndex        =   0
         Top             =   225
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar P1 
         Height          =   165
         Left            =   75
         TabIndex        =   2
         Top             =   1500
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione AFP"
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
         Index           =   2
         Left            =   225
         TabIndex        =   5
         Top             =   1110
         Width           =   1260
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
         Left            =   225
         TabIndex        =   4
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Mes Proceso"
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
         Left            =   225
         TabIndex        =   3
         Top             =   660
         Width           =   2010
      End
   End
End
Attribute VB_Name = "FrmRptAportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_ReportesAfp As ADODB.Recordset
Dim s_MesSeleccion As String
Public s_MesSeleccionRpt As String
Dim s_Codigo_Afp As String
Dim i_TotalRemuneracionAseg As Double
Public i_TotalAporteObligatorio As Double
Dim i_TotalComisionRA As Double
Dim i_TotalPrimaSeguro As Double
Public i_TotalRetencionRetribucion As Double
Dim r_FactorAporObligatorio As Single
Dim r_FactorPrimaseguro As Single
Dim r_FactorComisionRA As Single
Dim r_CantidadTope As Single
Dim s_CodTrabajador As String
Dim r_CantidadSuma As Single
Sub Llena_Barra()
    Dim i_Contador As Integer
    P1.Min = 1: P1.Max = 10
    For i_Contador = 1 To 10: P1.Value = i_Contador: Next i_Contador
End Sub
Sub Captura_Mes_Seleccionado()
    Select Case Cmbmes.Text
        Case "ENERO": s_MesSeleccion = "01": s_MesSeleccionRpt = "01"
        Case "FEBRERO": s_MesSeleccion = "02": s_MesSeleccionRpt = "02"
        Case "MARZO": s_MesSeleccion = "03": s_MesSeleccionRpt = "03"
        Case "ABRIL": s_MesSeleccion = "04": s_MesSeleccionRpt = "04"
        Case "MAYO": s_MesSeleccion = "05": s_MesSeleccionRpt = "05"
        Case "JUNIO": s_MesSeleccion = "06": s_MesSeleccionRpt = "06"
        Case "JULIO": s_MesSeleccion = "07": s_MesSeleccionRpt = "07"
        Case "AGOSTO": s_MesSeleccion = "08": s_MesSeleccionRpt = "08"
        Case "SETIEMBRE": s_MesSeleccion = "09": s_MesSeleccionRpt = "09"
        Case "OCTUBRE": s_MesSeleccion = "10": s_MesSeleccionRpt = "10"
        Case "NOVIEMBRE": s_MesSeleccion = "11": s_MesSeleccionRpt = "11"
        Case "DICIEMBRE": s_MesSeleccion = "12": s_MesSeleccionRpt = "12"
    End Select
End Sub
Sub Llena_Informacion_Adm_Pensiones()
    Call Recupera_Descripcion_Administradoras_Pensiones
    Set rs_ReportesAfp = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_ReportesAfp.EOF = False Then
        rs_ReportesAfp.MoveFirst
        Do While Not rs_ReportesAfp.EOF
            Cmbafp.AddItem rs_ReportesAfp!DESCRIP
            rs_ReportesAfp.MoveNext
        Loop
        Set rs_ReportesAfp = Nothing
    End If
End Sub
Sub Codigo_Afp_Seleccionado()
    Call Recupera_Codigo_Administradora_Pensiones(Cmbafp.Text)
    Set rs_ReportesAfp = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    s_Codigo_Afp = rs_ReportesAfp!cod_maestro2
    Set rs_ReportesAfp = Nothing
End Sub
Sub Monto_Total_Remuneracion_Asegurable()
    Call Recuperar_Monto_Total_Remuneracion_Asegurable(wcia, TxtAño, s_MesSeleccion, s_Codigo_Afp)
    Set rs_ReportesAfp = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_ReportesAfp.EOF = False Then
        If IsNull(rs_ReportesAfp!total_remaseg) Then
            i_TotalRemuneracionAseg = 0
        Else
            i_TotalRemuneracionAseg = rs_ReportesAfp!total_remaseg
        End If
    Else
        i_TotalRemuneracionAseg = 0
    End If
    Set rs_ReportesAfp = Nothing
End Sub
Sub Recupera_Factores()
    Call Recupera_Factores_Operacionales(wcia, TxtAño, s_MesSeleccion, s_Codigo_Afp)
    Set rs_ReportesAfp = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    r_FactorAporObligatorio = rs_ReportesAfp!Fac_AporOblig
    r_FactorPrimaseguro = rs_ReportesAfp!Fac_PrimaSeg
    r_FactorComisionRA = rs_ReportesAfp!Fac_ComRA
    r_CantidadTope = rs_ReportesAfp!tope
    Set rs_ReportesAfp = Nothing
End Sub
Sub Calculamos_Montos_Totales()
    i_TotalAporteObligatorio = (i_TotalRemuneracionAseg * r_FactorAporObligatorio) / 100
    i_TotalComisionRA = (i_TotalRemuneracionAseg * r_FactorComisionRA) / 100
    Call Recupera_Informacion_Calculo_Prima_Seguros(wcia, s_MesSeleccion, TxtAño, s_Codigo_Afp)
    Set rs_ReportesAfp = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_ReportesAfp.EOF = False Then
        rs_ReportesAfp.MoveFirst
        r_CantidadSuma = 0: i_TotalPrimaSeguro = 0
        s_CodTrabajador = rs_ReportesAfp!PlaCod
        Do While Not rs_ReportesAfp.EOF
            If s_CodTrabajador = rs_ReportesAfp!PlaCod Then
                r_CantidadSuma = r_CantidadSuma + rs_ReportesAfp!total_remaseg
            Else
                s_CodTrabajador = rs_ReportesAfp!PlaCod
                Call Calcula_Prima_Seguro
                r_CantidadSuma = rs_ReportesAfp!total_remaseg
            End If
            rs_ReportesAfp.MoveNext
        Loop
    End If
    Call Calcula_Prima_Seguro
    Set rs_ReportesAfp = Nothing
    i_TotalRetencionRetribucion = i_TotalComisionRA + i_TotalPrimaSeguro
End Sub
Sub Proceso_Ejecuta_Reporte()
    Call Codigo_Afp_Seleccionado
    Call Captura_Mes_Seleccionado
    Call Monto_Total_Remuneracion_Asegurable
    Call Recupera_Factores
    Call Calculamos_Montos_Totales
    Call Verifica_Existencia_Reporte3
    Call Crear_Plantilla_Reportes_Afp(wcia, s_MesSeleccion, TxtAño)
    Call Enviar_Montos_Aportaciones_Asegurables(i_TotalRemuneracionAseg, i_TotalAporteObligatorio, _
    i_TotalAporteObligatorio, i_TotalPrimaSeguro, i_TotalComisionRA, i_TotalRetencionRetribucion)
    Select Case Cmbafp.Text
        Case "PROFUTURO": i_Direccion_Reportes = 6
        Case "PRIMA-AFP": i_Direccion_Reportes = 3
        Case "HORIZONTE": i_Direccion_Reportes = 4
        Case "INTEGRA": i_Direccion_Reportes = 5
    End Select
End Sub
Sub Calcula_Prima_Seguro()
    If r_CantidadSuma < r_CantidadTope Then
        i_TotalPrimaSeguro = i_TotalPrimaSeguro + ((r_CantidadSuma * r_FactorPrimaseguro) / 100)
    Else
        i_TotalPrimaSeguro = i_TotalPrimaSeguro + ((r_CantidadTope * r_FactorPrimaseguro) / 100)
    End If
End Sub

Private Sub Form_Load()
    Call Llena_Informacion_Adm_Pensiones
    Call Llena_Barra
End Sub

