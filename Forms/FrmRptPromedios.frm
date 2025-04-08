VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form FrmRptPromedios 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Detalle Promedios"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4185
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "Procesar"
      Height          =   690
      Left            =   4275
      TabIndex        =   8
      Top             =   450
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1665
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3990
      Begin VB.TextBox TxtAño 
         Height          =   285
         Left            =   2325
         TabIndex        =   3
         Top             =   225
         Width           =   1515
      End
      Begin VB.ComboBox CmbMes 
         Height          =   315
         ItemData        =   "FrmRptPromedios.frx":0000
         Left            =   2325
         List            =   "FrmRptPromedios.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1515
      End
      Begin VB.ComboBox CmbTrabTipo 
         Height          =   315
         ItemData        =   "FrmRptPromedios.frx":0090
         Left            =   2325
         List            =   "FrmRptPromedios.frx":0092
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   975
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar P1 
         Height          =   165
         Left            =   75
         TabIndex        =   4
         Top             =   1425
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
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
         TabIndex        =   7
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesar al Mes"
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
         TabIndex        =   6
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
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
         Left            =   150
         TabIndex        =   5
         Top             =   1035
         Width           =   1350
      End
   End
End
Attribute VB_Name = "FrmRptPromedios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_RptPromedios As ADODB.Recordset
Dim rs_RptPromedios2 As ADODB.Recordset
Dim s_TipoTrabajador As String
Dim s_ConcaPromedios As String
Dim s_ConcaPromedios2 As String
Dim s_Mes_Seleccionado As Integer
Dim s_Mes_Seleccionado2 As String
Dim i_Parametro2 As Integer
Dim r_Monto_Procesar As Single
Private Sub Form_Load()
    Call Llena_Tipo_Trabajadores
    Call Llena_Barra
End Sub
Sub Llena_Tipo_Trabajadores()
    Call Recupera_Tipos_Trabajadores
    Set rs_RptPromedios = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    Do While Not rs_RptPromedios.EOF
        CmbTrabTipo.AddItem rs_RptPromedios!DESCRIP
        rs_RptPromedios.MoveNext
    Loop
    Set rs_RptPromedios = Nothing
End Sub
Sub Llena_Barra()
    Dim i_Contador As Integer
    P1.Min = 1: P1.Max = 10
    For i_Contador = 1 To 10: P1.Value = i_Contador: Next i_Contador
End Sub
Sub Recupera_tipo_trabajador()
    Select Case CmbTrabTipo
        Case "OBRERO": s_TipoTrabajador = "02"
        Case "EMPLEADO": s_TipoTrabajador = "01"
    End Select
End Sub
Sub Concatena_Informacion()
    Dim i_Contador_Letras As Integer
    Dim i_Contador_Letras2 As Integer
    s_ConcaPromedios = "": s_ConcaPromedios2 = ""
    Call Recupera_Informacion_Maestra1
    Set rs_RptPromedios = Reporte_Promedios.rs_Promedios_Pub
    Set Reporte_Promedios.rs_Promedios_Pub = Nothing
    s_ConcaPromedios = ""
    If rs_RptPromedios.EOF = False Then
        rs_RptPromedios.MoveFirst
        Do While Not rs_RptPromedios.EOF
            s_ConcaPromedios = s_ConcaPromedios & "p.i" & rs_RptPromedios!codinterno & ","
            s_ConcaPromedios2 = s_ConcaPromedios2 & "i" & rs_RptPromedios!codinterno & ","
            rs_RptPromedios.MoveNext
        Loop
        Set rs_RptPromedios = Nothing
        i_Contador_Letras = Len(s_ConcaPromedios)
        i_Contador_Letras2 = Len(s_ConcaPromedios2)
        s_ConcaPromedios = Mid(s_ConcaPromedios, 1, i_Contador_Letras - 1)
        s_ConcaPromedios2 = Mid(s_ConcaPromedios2, 1, i_Contador_Letras2 - 1)
    Else
    End If
End Sub
Sub Captura_Parametros_Meses()
    Select Case Cmbmes.Text
        Case "ENERO": s_Mes_Seleccionado = 1: Case "FEBRERO": s_Mes_Seleccionado = 2
        Case "MARZO": s_Mes_Seleccionado = 3: Case "ABRIL": s_Mes_Seleccionado = 4
        Case "MAYO": s_Mes_Seleccionado = 5: Case "JUNIO": s_Mes_Seleccionado = 6
        Case "JULIO": s_Mes_Seleccionado = 7: Case "AGOSTO": s_Mes_Seleccionado = 8
        Case "SETIEMBRE": s_Mes_Seleccionado = 9: Case "OCTUBRE": s_Mes_Seleccionado = 10
        Case "NOVIEMBRE": s_Mes_Seleccionado = 11: Case "DICIEMBRE": s_Mes_Seleccionado = 12
    End Select
    If s_Mes_Seleccionado < 7 Then
        i_Parametro2 = 0
    Else
    '23/01/2008
        i_Parametro2 = s_Mes_Seleccionado - 6  '7
    End If
End Sub
Sub Uniformizar_Informacion()

    Dim i_Contador_Vueltas As Integer
    Dim s_CodTrabajador_Uni As String
    Dim s_Columna_Sumar As String
    Dim s_MesProceso_Uni As Integer
    Dim i_Contador_Meses As Integer
    Dim i_Contador_Vueltas_Meses As Integer
    
    s_CodTrabajador_Uni = "": s_MesProceso_Uni = 0
    i_Contador_Vueltas_Meses = 0
    Call Recupera_Informacion_Promedios_Generales
    Set rs_RptPromedios = Reporte_Promedios.rs_Promedios_Pub
    Set Reporte_Promedios.rs_Promedios_Pub = Nothing
    If rs_RptPromedios.EOF = False Then
        rs_RptPromedios.MoveFirst
        Do While Not rs_RptPromedios.EOF
            If s_CodTrabajador_Uni <> rs_RptPromedios!PlaCod Or s_MesProceso_Uni <> Month(rs_RptPromedios!FechaProceso) Then
                Call Recupera_Informacion_Maestra1
                Set rs_RptPromedios2 = Reporte_Promedios.rs_Promedios_Pub
                Set Reporte_Promedios.rs_Promedios_Pub = Nothing
                rs_RptPromedios2.MoveFirst
                Do While Not rs_RptPromedios2.EOF
                    s_Columna_Sumar = "i" & rs_RptPromedios2!codinterno
                    Call Graba_Informacion_Uniforme(s_Columna_Sumar, rs_RptPromedios!PlaCod, Val(Month(rs_RptPromedios!FechaProceso)))
                    rs_RptPromedios2.MoveNext
                Loop
                Set rs_RptPromedios2 = Nothing
                s_CodTrabajador_Uni = rs_RptPromedios!PlaCod
                s_MesProceso_Uni = Val(Month(rs_RptPromedios!FechaProceso))
            End If
            rs_RptPromedios.MoveNext
        Loop
    End If
    Set rs_RptPromedios = Nothing
    For i_Contador_Meses = i_Parametro2 To Cmbmes.ListIndex
        i_Contador_Vueltas_Meses = i_Contador_Vueltas_Meses + 1
        Recupera_Mes_Proceso (i_Contador_Meses)
        Call Recupera_Promedios_Procesados_Por_Mes(wcia, s_Mes_Seleccionado2, s_ConcaPromedios2)
        Set rs_RptPromedios = Reporte_Promedios.rs_Promedios_Pub
        Set Reporte_Promedios.rs_Promedios_Pub = Nothing
        
        If rs_RptPromedios.EOF = False Then
            rs_RptPromedios.MoveFirst
            Do While rs_RptPromedios.EOF = False
                'If Trim(rs_RptPromedios!PLACOD) = "PO042" Then Stop
                Call Recupera_Informacion_Maestra1
                Set rs_RptPromedios2 = Reporte_Promedios.rs_Promedios_Pub
                Set Reporte_Promedios.rs_Promedios_Pub = Nothing
                rs_RptPromedios2.MoveFirst
                Do While Not rs_RptPromedios2.EOF
                
                    '*******este codigo hay que mejorarlo, esta asi por
                    'falta de tiempo***********************************
                    Select Case rs_RptPromedios2!codinterno
                        Case 10: r_Monto_Procesar = rs_RptPromedios!i10
                        Case 11: r_Monto_Procesar = rs_RptPromedios!i11
                        Case 16: r_Monto_Procesar = rs_RptPromedios!I16
                        Case 24: r_Monto_Procesar = rs_RptPromedios!I24
                        Case 25: r_Monto_Procesar = rs_RptPromedios!i25
                        Case 21: r_Monto_Procesar = rs_RptPromedios!i21
                        Case Else:: r_Monto_Procesar = 0
                    End Select
                    '***************************************************
                    
                    If Verifica_Existencia_Registro_Promedios(wcia, rs_RptPromedios!PlaCod, rs_RptPromedios2!Descripcion) = False Then
                        Call Graba_Nueva_Informacion_Promedios(rs_RptPromedios!cia, rs_RptPromedios!PlaCod, rs_RptPromedios!nombres, _
                        rs_RptPromedios2!Descripcion, r_Monto_Procesar, i_Contador_Vueltas_Meses)
                    Else
                        Call Edita_Informacion_Promedios(wcia, rs_RptPromedios!PlaCod, rs_RptPromedios!nombres, _
                        rs_RptPromedios2!Descripcion, r_Monto_Procesar, i_Contador_Vueltas_Meses)
                    End If
                    rs_RptPromedios2.MoveNext
                Loop
                rs_RptPromedios.MoveNext
            Loop
        End If
    Next i_Contador_Meses
End Sub
Sub Proceso_Central_Reporte_Promedios()
    Call Recupera_tipo_trabajador
    Call Genera_Tabla_Promedios_Maestra1(s_TipoTrabajador, wcia)
    Call Concatena_Informacion
    Call Captura_Parametros_Meses
    Call Crea_Tabla_Promedios_Temporal(TxtAño, s_Mes_Seleccionado, i_Parametro2, wcia, _
    s_TipoTrabajador, s_ConcaPromedios)
    Call Uniformizar_Informacion
    i_Direccion_Reportes = 8
    'RptAsistencias.Show
End Sub
Sub Recupera_Mes_Proceso(i_Mes As Integer)
    Select Case i_Mes
        Case 1: s_Mes_Seleccionado2 = "01"
        Case 2: s_Mes_Seleccionado2 = "02"
        Case 3: s_Mes_Seleccionado2 = "03"
        Case 4: s_Mes_Seleccionado2 = "04"
        Case 5: s_Mes_Seleccionado2 = "05"
        Case 6: s_Mes_Seleccionado2 = "06"
        Case 7: s_Mes_Seleccionado2 = "07"
        Case 8: s_Mes_Seleccionado2 = "08"
        Case 9: s_Mes_Seleccionado2 = "09"
        Case 10: s_Mes_Seleccionado2 = "10"
        Case 11: s_Mes_Seleccionado2 = "11"
        Case 12: s_Mes_Seleccionado2 = "12"
    End Select
End Sub

