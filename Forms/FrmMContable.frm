VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmMContable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Ingreso Contable Maestro «"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "FrmMContable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   690
      Left            =   3015
      TabIndex        =   4
      Top             =   150
      Width           =   3675
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         ItemData        =   "FrmMContable.frx":030A
         Left            =   150
         List            =   "FrmMContable.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   255
         Width           =   3345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Operación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   690
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   2700
      Begin VB.ComboBox CmbConOperacion 
         Height          =   315
         ItemData        =   "FrmMContable.frx":030E
         Left            =   150
         List            =   "FrmMContable.frx":031B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4590
      Left            =   75
      TabIndex        =   0
      Top             =   855
      Width           =   6615
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   300
         TabIndex        =   7
         Top             =   1500
         Visible         =   0   'False
         Width           =   6015
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "La Informacion de los Centros de Costo no es Uniforme"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   75
            TabIndex        =   8
            Top             =   225
            Width           =   5790
         End
      End
      Begin VB.TextBox Txtingreso 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   4425
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   1440
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid M1 
         Height          =   3990
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   7038
         _Version        =   393216
         ScrollBars      =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Top             =   4275
         Width           =   6315
      End
   End
End
Attribute VB_Name = "FrmMContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_Mcontable As ADODB.Recordset
Dim rs_Mcontable2 As ADODB.Recordset
Dim i_NumeroLinea As Integer
Dim i_NumeroLineaScroll As Integer
Dim i_TipoTeclado As Integer
Dim i_CapturaLineas As Integer
Dim s_CodTipo As String
Dim i_Longitud_Caja As String
Private Sub CmbConOperacion_Click()
    'If CmbConOperacion.Text <> "" And Cmbtipo.Text <> "" Then Call Ejecuta_Llenado_Grilla
    If CmbConOperacion.Text <> "" And Cmbtipo.ListIndex <> 0 Then Call Ejecuta_Llenado_Grilla
End Sub
Private Sub CmbTipo_Click()
'    Call Recupera_Codigo_Tipo_Trabajador(Cmbtipo.Text)
'    s_CodTipo = Crear_Plan_Contable.s_CodTipo_G
'    If CmbConOperacion.Text <> "" And Cmbtipo.Text <> "" Then Call Ejecuta_Llenado_Grilla
    s_CodTipo = Empty
    s_CodTipo = Trim(fc_CodigoComboBox(Cmbtipo, 2))
    If CmbConOperacion.Text <> "" And Cmbtipo.ListIndex <> 0 Then Call Ejecuta_Llenado_Grilla
End Sub
Private Sub Form_Activate()
    If Verifica_y_Captura_Longitud_Centro_Costo(wcia) = True Then
        'Txtingreso.MaxLength = 7 - Crear_Plan_Contable.i_Longitud_Centro_Costo
        i_Longitud_Caja = 8 - Crear_Plan_Contable.i_Longitud_Centro_Costo
    Else
        MsgBox "La Informacion de los Centros de Costo no es uniforme, Verifique por favor " & _
        "la Informacion Ingresada", vbCritical
        Frame4.Visible = True: Frame1.Enabled = False: Frame2.Enabled = False: Frame3.Enabled = False
        Exit Sub
    End If
    Label2.Caption = "Seleccione Tipo de Operacion y Tipo de Trabajador"
    Call Llena_Tipo_Trabajadores
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call Proceso_Integral
End Sub
Sub Ejecuta_Llenado_Grilla()
On Error GoTo MyErr
    M1.FixedCols = 1
    Select Case CmbConOperacion.Text
        Case "INGRESOS"
            Call Recupera_Conceptos_Ingresos(wcia, 2, s_CodTipo)
            Txtingreso.MaxLength = i_Longitud_Caja
            Label2.Caption = "Ingrese los " & i_Longitud_Caja & " Ultimos Digitos de la Cuenta Contable"
        Case "APORTACIONES"
            Call Recupera_Conceptos_Aportes(wcia, 2, s_CodTipo)
            Txtingreso.MaxLength = i_Longitud_Caja
            Label2.Caption = "Ingrese los " & i_Longitud_Caja & " Ultimos Digitos de la Cuenta Contable"
        Case "DEDUCCIONES"
            Call Recupera_Conceptos_Deduccion(wcia, 2, s_CodTipo)
            Txtingreso.MaxLength = 7
            Label2.Caption = "Ingrese Cuenta Contable de 7 Digitos "
    End Select
    Set M1.DataSource = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    With M1
      .ColWidth(0) = 300: .ColWidth(2) = 3000: .ColWidth(3) = 1300
    End With
    i_CapturaLineas = M1.Rows - 1
    Call Formato_Grilla
    
    M1.SetFocus: M1.Row = 1: M1.Col = 3: i_NumeroLinea = M1.Row
    Call Activa_Caja_Ingreso
MyErr:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub
Sub Llena_Tipo_Trabajadores()
'    Call Recupera_Tipos_Trabajadores
'    Set rs_Mcontable = Crear_Plan_Contable.rs_PlanCont_Pub
'    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
'    Do While Not rs_Mcontable.EOF
'        Cmbtipo.AddItem rs_Mcontable!descrip
'        rs_Mcontable.MoveNext
'    Loop
'    Set rs_Mcontable = Nothing
    Call Trae_Tipo_Trab(Cmbtipo)
End Sub
Sub Activa_Caja_Ingreso()
    Txtingreso.Visible = True: Txtingreso.Top = M1.CellTop + 220: Txtingreso.Left = M1.CellLeft + 160
    Txtingreso.Visible = True
    Txtingreso.Width = M1.CellWidth - 10: Txtingreso.Height = M1.CellHeight - 5: Txtingreso = M1.Text
    Txtingreso.BackColor = &HC0FFC0
    Txtingreso.SetFocus
    i_NumeroLineaScroll = M1.Row
End Sub
Sub Desactiva_Caja_Ingreso()
    M1.SetFocus: M1.Row = i_NumeroLinea: M1.Col = 3
    M1.Text = Txtingreso: Txtingreso.Visible = False: Txtingreso = "": i_NumeroLineaScroll = 0
End Sub
Sub Cambiar_Linea()
    Select Case i_TipoTeclado
        Case 1: If i_CapturaLineas <> M1.Row Then M1.Row = i_NumeroLinea + 1
        Case 2: If M1.Row <> 1 Then M1.Row = i_NumeroLinea - 1
        Case 3:
            If (M1.Row + 18) > i_CapturaLineas Then
                M1.Row = i_CapturaLineas
            Else
                M1.Row = M1.Row + 18
            End If
        Case 4
            If (M1.Row - 18) < 1 Then
                M1.Row = 1
            Else
                M1.Row = M1.Row - 18
            End If
    End Select
    i_NumeroLinea = M1.Row: Call Activa_Caja_Ingreso
End Sub
Private Sub M1_Click()
    Dim i_linea_boton As Integer
    If M1.Col <> 3 Then
        If Txtingreso.Visible = True Then
            Txtingreso.SetFocus
        End If
    Else
        i_linea_boton = M1.Row
        If Txtingreso.Visible = True Then Call Desactiva_Caja_Ingreso
        M1.Row = i_linea_boton
        i_NumeroLinea = M1.Row
        Call Activa_Caja_Ingreso
    End If
End Sub
Private Sub M1_Scroll()
    M1.SetFocus: M1.Row = i_NumeroLinea: M1.Col = 3
    If i_NumeroLineaScroll = M1.Row Then
        If Txtingreso.Visible = True Then
            M1.Text = Txtingreso
        End If
    End If
    Txtingreso.Visible = False: Txtingreso = ""
End Sub
Private Sub TxtIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40: i_TipoTeclado = 1: Call Desactiva_Caja_Ingreso: Call Cambiar_Linea
        Case 38: i_TipoTeclado = 2: Call Desactiva_Caja_Ingreso: Call Cambiar_Linea
        Case 34: i_TipoTeclado = 3: Call Desactiva_Caja_Ingreso: Call Cambiar_Linea
        Case 33: i_TipoTeclado = 4: Call Desactiva_Caja_Ingreso: Call Cambiar_Linea
        Case 13: i_TipoTeclado = 1: Call Desactiva_Caja_Ingreso: Call Cambiar_Linea
    End Select
End Sub
Sub Formato_Grilla()
    Dim i_Contador As Integer
    M1.FixedCols = 3: M1.SetFocus
    With M1
        For i_Contador = 1 To i_CapturaLineas
            .Row = i_Contador
            .Col = 1: .CellBackColor = &H80000005
            .Col = 2: .CellBackColor = &H80000005
        Next i_Contador
    End With
End Sub
Sub Graba_Informacion_Ingresada()
    Dim i_Contador_Filas As Integer
    Dim s_clave As String
    Dim s_Centro_Costo As String
    Dim s_CtaContable As String
    Call Desactiva_Caja_Ingreso
    Call Activa_Caja_Ingreso
    If Not ValCuenta Then Exit Sub
    With M1
        For i_Contador_Filas = 1 To .Rows - 1
            .Row = i_Contador_Filas
            .Col = 1: s_clave = .Text & s_CodTipo
            .Col = 3: s_CtaContable = .Text
            Select Case CmbConOperacion.Text
                Case "APORTACIONES": s_Centro_Costo = Mid(.Text, 1, 3)
                Case "DEDUCCIONES": s_Centro_Costo = ""
                Case "INGRESOS": s_Centro_Costo = Mid(.Text, 1, 3)
            End Select
            Call Graba_Informacion_Contable_Maestros2(s_clave, s_Centro_Costo, s_CtaContable, wcia)
        Next i_Contador_Filas
    End With
    MsgBox "Informacion Grabada Satisfactoriamente", vbInformation
End Sub
Private Function ValCuenta() As Boolean
    ValCuenta = False
    Dim I As Integer
    Dim texto As String
    With M1
        For I = 1 To .Rows - 1
            .Row = I
            .Col = 2: texto = .Text
            Select Case CmbConOperacion.Text
                Case "INGRESOS"
                    .Col = 3: If Len(Trim(.Text)) < 4 And Trim(.Text) <> Empty Then MsgBox "Debe de ingresar los 4 ultimos digitos de la cuenta, Verifique." & vbCrLf & "Concepto : " & texto, vbOKOnly + vbCritical, "Sistema": Exit Function
                Case "APORTACIONES"
                    .Col = 3: If Len(Trim(.Text)) < 4 And Trim(.Text) <> Empty Then MsgBox "Debe de ingresar los 4 ultimos digitos de la cuenta, Verifique." & vbCrLf & "Concepto : " & texto, vbOKOnly + vbCritical, "Sistema": Exit Function
                Case "DEDUCCIONES"
                    .Col = 3: If Len(Trim(.Text)) < 7 And Trim(.Text) <> Empty Then MsgBox "Debe de ingresar los 7 digitos de la cuenta, Verifique." & vbCrLf & "Concepto : " & texto, vbOKOnly + vbCritical, "Sistema": Exit Function
            End Select
        Next I
    End With
    ValCuenta = True
End Function


