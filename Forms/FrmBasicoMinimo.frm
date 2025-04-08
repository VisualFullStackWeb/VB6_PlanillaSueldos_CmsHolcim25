VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBasicoMinimo 
   Caption         =   "Basico Minimo"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBasicoMinimo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   5895
   Begin VB.TextBox TxtMinimo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   60
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.TextBox TxtMeses 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.ComboBox Cmbtipotrabajador 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmBasicoMinimo.frx":030A
      Left            =   120
      List            =   "FrmBasicoMinimo.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   975
      Width           =   600
      BackColor       =   12632256
      PicturePosition =   327683
      Size            =   "1058;873"
      Picture         =   "FrmBasicoMinimo.frx":030E
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      Height          =   195
      Left            =   3960
      TabIndex        =   8
      Top             =   765
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meses"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   765
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   765
      Width           =   1365
   End
End
Attribute VB_Name = "FrmBasicoMinimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmbcia_Click()
Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
End Sub

Private Sub Cmbtipotrabajador_Click()
Carga_Datos
End Sub

Private Sub CommandButton2_Click()
If Not IsNumeric(TxtMeses.Text) Then
   MsgBox "Ingrese Correctamente Meses", vbInformation: Exit Sub
End If
If Not IsNumeric(TxtMinimo.Text) Then
   MsgBox "Ingrese Correctamente Importe", vbInformation: Exit Sub
End If

Sql$ = "update Pla_Basico_Minimo set status='*' where cia='" & wcia & "' and tipotrab='" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "' and status<>'*'"
cn.Execute Sql$
Sql$ = "Insert into Pla_Basico_Minimo values ('" & wcia & "','" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "'," & CCur(TxtMeses.Text) & "," & CCur(TxtMinimo.Text) & ",'','" & wuser & "',getdate())"
cn.Execute Sql$
MsgBox "Grabación Satisfactoria", vbInformation
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 6015: Me.Height = 2220
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub Carga_Datos()
TxtMeses.Text = "": TxtMinimo.Text = ""
Sql$ = "Select * From Pla_Basico_Minimo where cia='" & wcia & "' and tipotrab='" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   TxtMeses.Text = rs!meses: TxtMinimo.Text = rs!importe
End If
rs.Close
End Sub
