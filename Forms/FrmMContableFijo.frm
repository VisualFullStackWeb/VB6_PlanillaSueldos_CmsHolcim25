VERSION 5.00
Begin VB.Form FrmMContableFijo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Ingreso Contable Fijo - Cuentas Netos «"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "FrmMContableFijo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4545
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   4260
      Begin VB.ComboBox CboTipo_Trab 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmMContableFijo.frx":030A
         Left            =   1455
         List            =   "FrmMContableFijo.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   2670
      End
      Begin VB.TextBox TxtIngreso 
         Height          =   315
         Index           =   3
         Left            =   2745
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1815
         Width           =   1365
      End
      Begin VB.TextBox TxtIngreso 
         Height          =   315
         Index           =   2
         Left            =   2745
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox TxtIngreso 
         Height          =   315
         Index           =   1
         Left            =   2745
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1065
         Width           =   1365
      End
      Begin VB.TextBox TxtIngreso 
         Height          =   315
         Index           =   0
         Left            =   2745
         MaxLength       =   8
         TabIndex        =   4
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo. Trabajador"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   10
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo de Servicion (CTS)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   7
         Top             =   1815
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Gratificaciones"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   3
         Top             =   1440
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Vacaciones"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   2
         Top             =   1065
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Normal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   690
         Width           =   1725
      End
   End
End
Attribute VB_Name = "FrmMContableFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_contablefijo As ADODB.Recordset
Dim mTipo_Trab As String

Sub Grabar_Informacion_Nueva()
    If CboTipo_Trab.ListIndex = 0 Then Exit Sub
    Dim i_Contador As Integer
    Dim s_clave As String
    Dim s_Cta_Contable As String
    For i_Contador = 1 To 4
        Select Case i_Contador
            Case 1: s_clave = "001": s_Cta_Contable = Txtingreso(0)
            Case 2: s_clave = "002": s_Cta_Contable = Txtingreso(1)
            Case 3: s_clave = "003": s_Cta_Contable = Txtingreso(2)
            Case 4: s_clave = "006": s_Cta_Contable = Txtingreso(3)
        End Select
        Call Grabar_Informacion_Nueva_Fijo(s_Cta_Contable, s_clave, wcia, mTipo_Trab)
    Next i_Contador
    MsgBox "Informacion Grabada Satisfactoriamente", vbInformation
End Sub

Private Sub CboTipo_Trab_Click()
    mTipo_Trab = Empty
    mTipo_Trab = Trim(fc_CodigoComboBox(CboTipo_Trab, 2))
    Call LimpiarTexto
    Call Llena_Informacion
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call Proceso_Integral
    Call Trae_Tipo_Trab(CboTipo_Trab)
End Sub

Private Sub TxtIngreso_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            Select Case index
                Case Is < 3: Txtingreso(index + 1).SetFocus
                Case 3: Txtingreso(0).SetFocus
            End Select
    End Select
End Sub

Private Sub Llena_Informacion()
    If CboTipo_Trab.ListIndex = 0 Then Exit Sub
    Dim i_Contador As Integer
    Call Recupera_Informacion_Contable_Fijo(mTipo_Trab)
    Set rs_contablefijo = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    If Not rs_contablefijo.EOF Then
        rs_contablefijo.MoveFirst
        Do While Not rs_contablefijo.EOF
            Select Case rs_contablefijo!cmaf_clave
                Case "001": Txtingreso(0) = rs_contablefijo!cmaf_ctacontable
                Case "002": Txtingreso(1) = rs_contablefijo!cmaf_ctacontable
                Case "003": Txtingreso(2) = rs_contablefijo!cmaf_ctacontable
                Case "006": Txtingreso(3) = rs_contablefijo!cmaf_ctacontable
            End Select
            rs_contablefijo.MoveNext
        Loop
    End If
    Set rs_contablefijo = Nothing
End Sub

Private Sub LimpiarTexto()
    Txtingreso(0).Text = Empty
    Txtingreso(1).Text = Empty
    Txtingreso(2).Text = Empty
    Txtingreso(3).Text = Empty
End Sub

