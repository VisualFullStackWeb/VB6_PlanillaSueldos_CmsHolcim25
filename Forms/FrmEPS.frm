VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmEPS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Entidad Prestadora de Salud «"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "FrmEPS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_cod 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   1215
      TabIndex        =   12
      Top             =   855
      Width           =   675
   End
   Begin VB.Frame frm_00 
      Height          =   630
      Left            =   3075
      TabIndex        =   5
      Top             =   780
      Width           =   1905
      Begin MSForms.CommandButton btn_salir 
         Height          =   375
         Left            =   1455
         TabIndex        =   10
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmEPS.frx":030A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_eliminar 
         Height          =   375
         Left            =   1110
         TabIndex        =   7
         Top             =   165
         Visible         =   0   'False
         Width           =   360
         VariousPropertyBits=   268435481
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmEPS.frx":08A4
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_cancelar 
         Height          =   375
         Left            =   765
         TabIndex        =   6
         Top             =   165
         Visible         =   0   'False
         Width           =   360
         VariousPropertyBits=   268435481
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmEPS.frx":0E3E
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_aceptar 
         Height          =   375
         Left            =   420
         TabIndex        =   8
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmEPS.frx":13D8
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_nuevo 
         Height          =   375
         Left            =   75
         TabIndex        =   9
         Top             =   165
         Visible         =   0   'False
         Width           =   360
         VariousPropertyBits=   268435481
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmEPS.frx":1972
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   105
      Width           =   3765
   End
   Begin VB.TextBox txt_Aportacion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1215
      TabIndex        =   1
      Top             =   465
      Width           =   1575
   End
   Begin VB.TextBox txt_descrip 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   105
      Width           =   3765
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id Sunat"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   855
      Width           =   600
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E.P.S"
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   4
      Top             =   105
      Width           =   405
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aportacion %"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   450
      Width           =   930
   End
End
Attribute VB_Name = "FrmEPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim pCodigo As String

Private Sub Load_Info()
    Cadena = "SP_TRAE_ENTIDAD_PRESTADORA_SALUD"
    Set rs = OpenRecordset(Cadena, cn)
    
    If Not rs.EOF Then
        Cbo.Clear
        Do While Not rs.EOF
            Cbo.AddItem Trim(rs!DESCRIP)
            Cbo.ItemData(Cbo.NewIndex) = Trim(rs!cod_maestro2)
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Load_EPS()
    With rs
        .MoveFirst
         Do While Not .EOF
            If Trim(!cod_maestro2) = fc_CodigoComboBox(Cbo, 2) Then
                txt_Aportacion.Text = Trim(!importe)
                txt_cod.Text = Trim(!CODSUNAT)
                Exit Do
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub Clear()
    Cbo.Clear
    Cbo.Visible = True
    txt_Aportacion.Text = Empty
    txt_cod.Text = Empty
    txt_descrip.Text = Empty
    txt_descrip.Visible = False
End Sub

Private Sub btn_aceptar_Click()
    Call Aceptar
End Sub

Private Sub btn_salir_Click()
    Unload Me
End Sub

Private Sub Cbo_Click()
    Call Load_EPS
    pCodigo = Empty
    pCodigo = fc_CodigoComboBox(Cbo, 2)
End Sub

Public Sub Nuevo()
    Call Clear
    Cbo.Visible = False
    txt_descrip.Visible = True
    On Error Resume Next
    txt_descrip.SetFocus
End Sub

Public Sub Aceptar()
    If pCodigo = "" Then Exit Sub
    Cadena = "SP_MANT_EPS '" & pCodigo & "', " & Val(txt_Aportacion.Text) & ""
    If EXEC_SQL(Cadena, cn) Then
        MsgBox "Registro grabado satisfactoriamente.", vbInformation + vbOKOnly, "Sistema"
        Call Form_Load
    Else
        MsgBox "Error al grabar el registro." & "Se cerrará el formulario.", vbCritical + vbOKOnly, "Sistema"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call Clear
    Call Load_Info
End Sub
