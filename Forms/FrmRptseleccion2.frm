VERSION 5.00
Begin VB.Form FrmRptseleccion2 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Lista Trabajadores"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3525
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   1065
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3315
      Begin VB.ComboBox CmbTip 
         Height          =   315
         ItemData        =   "FrmRptseleccion2.frx":0000
         Left            =   225
         List            =   "FrmRptseleccion2.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   525
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Tipo Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2955
      End
   End
End
Attribute VB_Name = "FrmRptseleccion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_RptSelecc As ADODB.Recordset
Dim rs_RptSelecc2 As ADODB.Recordset

Private Sub Form_Load()
    Call Llena_Tipo_Trabajador
End Sub
Sub Llena_Tipo_Trabajador()
    Call Recupera_Tipos_Trabajadores
    Set rs_RptSelecc = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    Do While Not rs_RptSelecc.EOF
        CmbTip.AddItem rs_RptSelecc!DESCRIP
        rs_RptSelecc.MoveNext
    Loop
    Set rs_RptSelecc = Nothing
End Sub
Sub Proceso_Reporte_Dos()
    Call Generar_Proceso_Tabla_Reporte_Lista_Trabajadores(wcia)
    i_Direccion_Reportes = 2
    'RptAsistencias.Show
End Sub

