VERSION 5.00
Begin VB.Form FrmCtaProvision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Cuentas Provisión «"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "FrmCtaProvision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4575
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4410
      Begin VB.TextBox TxtProvision 
         Height          =   285
         Left            =   1410
         TabIndex        =   8
         Top             =   1350
         Width           =   1275
      End
      Begin VB.TextBox TxtCcosto 
         Height          =   315
         Left            =   1410
         TabIndex        =   6
         Top             =   975
         Width           =   1275
      End
      Begin VB.ComboBox CmbTrabTipo 
         Height          =   315
         ItemData        =   "FrmCtaProvision.frx":030A
         Left            =   1410
         List            =   "FrmCtaProvision.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   2895
      End
      Begin VB.ComboBox CmbBoleta 
         Height          =   315
         ItemData        =   "FrmCtaProvision.frx":030E
         Left            =   1410
         List            =   "FrmCtaProvision.frx":031B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   150
         TabIndex        =   9
         Top             =   1725
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta Provisión"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   1350
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta Centro Costo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   975
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Provisión"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   3
         Top             =   600
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FrmCtaProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_ctaProvision As ADODB.Recordset
Dim i_Longitud_Caja As Integer
Dim s_TipoTrabajador As String

Private Sub CmbBoleta_Click()
    If CmbBoleta.Text <> "" And s_TipoTrabajador <> "" Then
        Call Recupera_Informacion_Ingresada
    End If
End Sub

Private Sub CmbTrabTipo_Click()
    s_TipoTrabajador = Empty
    s_TipoTrabajador = Trim(fc_CodigoComboBox(CmbTrabTipo, 2))
    
    If CmbBoleta.Text <> "" And s_TipoTrabajador <> "" Then
        Call Recupera_Informacion_Ingresada
    End If
End Sub

Private Sub Form_Activate()
    If Verifica_y_Captura_Longitud_Centro_Costo(wcia) = True Then
        i_Longitud_Caja = 8 - Crear_Plan_Contable.i_Longitud_Centro_Costo
    Else
        MsgBox "La Informacion de los Centros de Costo no es uniforme, Verifique por favor " & _
        "la Informacion Ingresada", vbCritical
        Label2.Caption = "Centro de Costos No Uniforme": Frame1.Enabled = False
        Exit Sub
    End If
    TxtCcosto.MaxLength = i_Longitud_Caja
End Sub

Private Sub Form_Load()
'    Call Llena_Tipo_Trabajadores
    Me.Top = 0
    Me.Left = 0
    Call Trae_Tipo_Trab(CmbTrabTipo)
End Sub

'Sub Llena_Tipo_Trabajadores()
'    Call Recupera_Tipos_Trabajadores
'    Set rs_ctaProvision = Crear_Plan_Contable.rs_PlanCont_Pub
'    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
'    Do While Not rs_ctaProvision.EOF
'        CmbTrabTipo.AddItem rs_ctaProvision!descrip
'        rs_ctaProvision.MoveNext
'    Loop
'    Set rs_ctaProvision = Nothing
'End Sub

Private Sub TxtCcosto_GotFocus()
    Label2.Caption = "Ingrese los " & i_Longitud_Caja & " Ultimos Digitos"
End Sub

Private Sub TxtCcosto_LostFocus()
    Label2.Caption = ""
End Sub

Sub Graba_Informacion_Ingresada()
    'Call Recupera_tipo_trabajador
    If CmbTrabTipo.ListIndex = 0 Then Exit Sub
    If Verifica_Existencia_registro_Parametros(wcia, s_TipoTrabajador, CmbBoleta.Text) = False Then
        Call Graba_Informacion_Asiento_Provision_Parametros(wcia, s_TipoTrabajador, CmbBoleta.Text, _
        TxtProvision, TxtCcosto)
    Else
        Call Edita_Informacion_Asiento_Provision_Parametros(wcia, s_TipoTrabajador, CmbBoleta.Text, _
        TxtProvision, TxtCcosto)
    End If
    MsgBox "Grabacion Efectuada Exitosamente"
End Sub

'Sub Recupera_tipo_trabajador()
'    Select Case CmbTrabTipo
'        Case "OBRERO": s_TipoTrabajador = "02"
'        Case "EMPLEADO": s_TipoTrabajador = "01"
'    End Select
'End Sub

Sub Recupera_Informacion_Ingresada()
    'Call Recupera_tipo_trabajador
    If CmbTrabTipo.ListIndex = 0 Then Exit Sub
    If Verifica_Existencia_registro_Parametros(wcia, s_TipoTrabajador, CmbBoleta.Text) = True Then
        Call Recupera_Informacion_Parametros_Asiento_Provision(wcia, s_TipoTrabajador, _
        CmbBoleta.Text)
        Set rs_ctaProvision = Reportes_Centrales.rs_RptCentrales_pub
        Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
        If rs_ctaProvision.EOF = False Then
            rs_ctaProvision.MoveFirst
            TxtProvision = rs_ctaProvision!CtaProvision
            TxtCcosto = rs_ctaProvision!ctacentrocosto
        End If
        Set rs_ctaProvision = Nothing
    Else
        TxtProvision = "": TxtCcosto = ""
    End If
End Sub

Private Sub TxtProvision_GotFocus()
    Label2.Caption = "Cuenta de Provisión de " & CmbBoleta.Text
End Sub

Private Sub TxtProvision_LostFocus()
    Label2.Caption = ""
End Sub
