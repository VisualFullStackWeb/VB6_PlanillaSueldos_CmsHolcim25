VERSION 5.00
Begin VB.Form FrmMContAsoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Configuración Adicional Cuentas Contables «"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "FrmMContAsoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4485
      Begin VB.TextBox Txtingreso 
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   8
         TabIndex        =   5
         Top             =   975
         Width           =   1890
      End
      Begin VB.TextBox Txtingreso 
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   8
         TabIndex        =   4
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empleados"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   3
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obreros"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   2
         Top             =   615
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deducciones  Asociadas al Seguro Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   225
         TabIndex        =   1
         Top             =   225
         Width           =   3360
      End
   End
End
Attribute VB_Name = "FrmMContAsoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_ContaAsociadas As ADODB.Recordset
Dim mTipo_Trab As String

Sub Grabar_Informacion_Nueva()
    Dim i_Contador As Integer
    Dim s_clave As String
    Dim s_Cta_Contable As String
    For i_Contador = 1 To 2
        Select Case i_Contador
            Case 1: s_clave = "004": s_Cta_Contable = Txtingreso(0)
            Case 2: s_clave = "005": s_Cta_Contable = Txtingreso(1)
        End Select
        Call Grabar_Informacion_Nueva_Fijo(s_Cta_Contable, s_clave, wcia)
    Next i_Contador
    MsgBox "Informacion Grabada Satisfactoriamente", vbInformation
End Sub
Private Sub Form_Activate()
    Call Llena_Informacion
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call Proceso_Integral
End Sub
Sub Llena_Informacion()
    Dim i_Contador As Integer
    Call Recupera_Informacion_Contable_Fijo
    Set rs_ContaAsociadas = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    If Not rs_ContaAsociadas.EOF Then
    rs_ContaAsociadas.MoveFirst
    Do While Not rs_ContaAsociadas.EOF
        Select Case rs_ContaAsociadas!cmaf_clave
            Case "004": Txtingreso(0) = rs_ContaAsociadas!cmaf_ctacontable
            Case "005": Txtingreso(1) = rs_ContaAsociadas!cmaf_ctacontable
            'Case "003": Txtingreso(2) = rs_contablefijo!cmaf_ctacontable
        End Select
        rs_ContaAsociadas.MoveNext
    Loop
    End If
    Set rs_ContaAsociadas = Nothing
End Sub

