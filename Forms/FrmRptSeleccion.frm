VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptSeleccion 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Asistencia Diaria"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6150
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1890
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   5940
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3975
         TabIndex        =   9
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   39325
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   2025
         TabIndex        =   8
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   39325
      End
      Begin VB.TextBox TxtPagina 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4425
         TabIndex        =   6
         Top             =   375
         Width           =   765
      End
      Begin VB.TextBox TxtHorario 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   1125
         Width           =   5640
      End
      Begin VB.ComboBox CmbTip 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   2490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   7
         Top             =   1575
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Pagina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   2925
         TabIndex        =   5
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horario de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   3
         Top             =   885
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   0
         X2              =   5925
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Tipo Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2370
      End
   End
End
Attribute VB_Name = "FrmRptSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_RptSeleccion As ADODB.Recordset
Dim s_TipTrajadorRpt As String
Dim i_Mes1 As Integer
Dim i_Mes2 As Integer
Dim i_Dia1 As Integer
Dim I_dia2 As Integer
Dim i_Numero_Pagina As Integer
Private Sub CmbTip_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13: TxtPagina.SetFocus
    End Select
End Sub
Private Sub Form_Load()
    Call Llena_Tipo_Trabajador
    Call Proceso_Horario_Fecha
End Sub
Sub Llena_Tipo_Trabajador()
    Call Recupera_Tipos_Trabajadores
    Set rs_RptSeleccion = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    Do While Not rs_RptSeleccion.EOF
        CmbTip.AddItem rs_RptSeleccion!DESCRIP
        rs_RptSeleccion.MoveNext
    Loop
    Set rs_RptSeleccion = Nothing
End Sub
Sub Procesar_Reporte()
    Dim i_Contador As Integer
    Dim i_Contador_Dias As Integer
    Dim s_Nombre As String
    
    i_Mes1 = Val(Month(DTPicker1.Value))
    i_Mes2 = Val(Month(DTPicker2.Value))
    
    If i_Mes1 <> i_Mes2 Then
        MsgBox "El Rango de Fecha debe ser del mismo mes de Seleccion", vbInformation
        Exit Sub
    End If
    
    i_Dia1 = Val(Day(DTPicker1.Value))
    I_dia2 = Val(Day(DTPicker2.Value))
    
    i_Contador = 0
    Call Recupera_Informacion_RegistroExistencia(wcia)
    Set rs_RptSeleccion = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_RptSeleccion.EOF = False Then
        Call Inserta_Inforamcion_Nueva(wcia, TxtHorario, Val(TxtPagina))
    Else
        Call Graba_Informacion_Nueva_Registro_Asistencia(wcia, TxtHorario, Val(TxtPagina))
    End If
    Call Recupera_tipo_trabajador
    
    i_Numero_Pagina = Val(TxtPagina)
    For i_Contador_Dias = i_Dia1 To I_dia2
        Call Genera_Informacion_Reporte_Asistencia(wcia, s_TipTrajadorRpt, _
        i_Contador_Dias, i_Numero_Pagina)
        i_Numero_Pagina = i_Numero_Pagina + 1
    Next i_Contador_Dias
    
    For i_Contador_Dias = i_Dia1 To I_dia2
        i_Contador = 0
        Call Recupera_Nombres_Reporte(i_Contador_Dias)
        Set rs_RptSeleccion = Reportes_Centrales.rs_RptCentrales_pub
        Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
        If rs_RptSeleccion.EOF = False Then
            rs_RptSeleccion.MoveFirst
            Do While Not rs_RptSeleccion.EOF
                i_Contador = i_Contador + 1
                s_Nombre = i_Contador & ".- " & rs_RptSeleccion!nombres
                Call Ingresa_Nueva_Informacion(s_Nombre, rs_RptSeleccion!nombres, _
                i_Contador, i_Contador_Dias)
                rs_RptSeleccion.MoveNext
            Loop
        End If
        Set rs_RptSeleccion = Nothing
    Next i_Contador_Dias

    Call Inserta_Inforamcion_Nueva(wcia, TxtHorario, i_Numero_Pagina - 1)
    
    i_Direccion_Reportes = 1
    'RptAsistencias.tdFecIni = DTPicker1.Value
    'RptAsistencias.Show
End Sub
Sub Recupera_tipo_trabajador()
    Select Case CmbTip.Text
        Case "OBRERO": s_TipTrajadorRpt = "02"
        Case "EMPLEADO": s_TipTrajadorRpt = "01"
    End Select
End Sub
Sub Proceso_Horario_Fecha()
    If Verifica_Existencia_Tabla_RegistroExistencia(wcia) = True Then
        Set rs_RptSeleccion = Reportes_Centrales.rs_RptCentrales_pub
        Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
        If rs_RptSeleccion.EOF = False Then
            TxtPagina = rs_RptSeleccion!Pagina + 1: TxtHorario = rs_RptSeleccion!Horario
        End If
        Set rs_RptSeleccion = Nothing
    Else
        Call Crear_Tabla_RegistroAsistencia(wcia)
        TxtPagina = 1
    End If
End Sub
Private Sub TxtHorario_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13: CmbTip.SetFocus
    End Select
End Sub
Private Sub TxtPagina_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13: TxtHorario.SetFocus
    End Select
End Sub
