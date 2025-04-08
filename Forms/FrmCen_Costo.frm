VERSION 5.00
Begin VB.Form FrmCen_Costo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Centro de Costos «"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "FrmCen_Costo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSenati 
      Caption         =   "Senati"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox Cmbtipotrabajador 
      Height          =   315
      ItemData        =   "FrmCen_Costo.frx":030A
      Left            =   1440
      List            =   "FrmCen_Costo.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton btn_Delete 
      Caption         =   "&Eliminar"
      Height          =   450
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txt_descrip 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton btn_Ok 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   3480
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton btn_Add 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   1785
      Width           =   1215
   End
   Begin VB.TextBox txt_Cen_Costo 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txt_Cta 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cta.CC"
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   510
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CC / Area"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   705
   End
End
Attribute VB_Name = "FrmCen_Costo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim Cadena As String
Dim bNuevo As Boolean

Private Sub btn_Add_Click()
    Call Operacion
End Sub

Private Sub btn_Delete_Click()
If bNuevo Then Exit Sub
If MsgBox("Desea grabar el registro?", vbQuestion + vbYesNo, "Sistema") = vbYes Then
    Cadena = "SP_DELETE_CEN_COSTO '" & wcia & "','" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "','" & fc_CodigoComboBox(Cbo, 2) & "'"
    If EXEC_SQL(Cadena, cn) Then
        MsgBox "Registro grabado satisfactoriamente.", vbInformation + vbOKOnly, "Sistema"
        Call Form_Load
    Else
        MsgBox "Error al grabar el registro." & "Se cerrará el formulario.", vbCritical + vbOKOnly, "Sistema"
        Unload Me
    End If
End If
End Sub

Private Sub btn_Ok_Click()
If MsgBox("Desea grabar el registro?", vbQuestion + vbYesNo, "Sistema") = vbYes Then
    If bNuevo Then
        If Len(Trim(txt_descrip.Text)) = 0 Then MsgBox "Ingrese la Descripción.": Exit Sub
        Cadena = "SP_MANT_CEN_COSTO '" & wcia & "', '" & fc_CodigoComboBox(Me.Cmbtipotrabajador, 2) & "', '', '" & Trim(txt_Cta.Text) & "', '" & Trim(txt_Cen_Costo.Text) & "','" & Trim(txt_descrip.Text) & "','" & chkSenati.Value & "'"
    Else
        Cadena = "SP_MANT_CEN_COSTO '" & wcia & "', '" & fc_CodigoComboBox(Me.Cmbtipotrabajador, 2) & "', '" & fc_CodigoComboBox(Cbo, 2) & "', '" & Trim(txt_Cta.Text) & "', '" & Trim(txt_Cen_Costo.Text) & "','','" & chkSenati.Value & "'"
    End If
    If EXEC_SQL(Cadena, cn) Then
        MsgBox "Registro grabado satisfactoriamente.", vbInformation + vbOKOnly, "Sistema"
        Call Form_Load
    Else
        MsgBox "Error al grabar el registro." & "Se cerrará el formulario.", vbCritical + vbOKOnly, "Sistema"
        Call Form_Load
    End If
End If
End Sub

Private Sub Cbo_Click()
    Call Load_Cta_CenCosto
End Sub

Private Sub Cmbtipotrabajador_Click()
Call Load_Info(fc_CodigoComboBox(Cmbtipotrabajador, 2))

End Sub

Private Sub Form_Load()
    Call Clear
    'Call Operacion
    Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
    If Cmbtipotrabajador.ListCount > 0 Then Cmbtipotrabajador.ListIndex = 0
       
    
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Load_Info(TipoTrab As String)
    Cadena = "SP_TRAE_INFO_CEN_COSTO '" & wcia & "','" & TipoTrab & "'"
    Set rs = OpenRecordset(Cadena, cn)
    
    If Not rs.EOF Then
        Cbo.Clear
        Do While Not rs.EOF
            Cbo.AddItem Trim(rs!DESCRIP_ARE)
            Cbo.ItemData(Cbo.NewIndex) = Trim(rs!COD_MAESTRO3)
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Load_Cta_CenCosto()

    With rs
        .MoveFirst
         Do While Not .EOF
            If Trim(!COD_MAESTRO3) = fc_CodigoComboBox(Cbo, 2) Then
                txt_Cta.Text = Trim(!CUENTA)
                txt_Cen_Costo.Text = Trim(!CENCOSTO)
                Me.chkSenati.Value = Val(!flag3)
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    
End Sub

Private Sub Clear()
    Cbo.Clear
    Cbo.Visible = True
    txt_Cta.Text = Empty
    txt_Cen_Costo.Text = Empty
    txt_descrip.Text = Empty
    chkSenati.Value = 0
    bNuevo = False
    txt_descrip.Visible = False
End Sub

Private Sub Operacion()
    If bNuevo Then
        btn_Add.Caption = "&Nuevo"
        Call Form_Load
        'Call Clear
    Else
        btn_Add.Caption = "&Cancelar"
        Call Clear
        Cbo.Visible = False
        txt_descrip.Visible = True
        On Error Resume Next
        txt_descrip.SetFocus
        bNuevo = True
    End If
End Sub
