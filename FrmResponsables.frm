VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Mantenimientos de Cargos «"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "FrmCargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtAbrev 
      Height          =   315
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   12
      Top             =   120
      Width           =   1605
   End
   Begin VB.TextBox txt_cod 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1155
      TabIndex        =   6
      Top             =   90
      Width           =   1575
   End
   Begin VB.Frame frm_00 
      Height          =   630
      Left            =   4680
      TabIndex        =   0
      Top             =   4710
      Width           =   1905
      Begin MSForms.CommandButton btn_salir 
         Height          =   375
         Left            =   1455
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmCargo.frx":030A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_eliminar 
         Height          =   375
         Left            =   1110
         TabIndex        =   2
         ToolTipText     =   "Eliminar"
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmCargo.frx":08A4
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_editar 
         Height          =   375
         Left            =   765
         TabIndex        =   3
         ToolTipText     =   "Editar"
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmCargo.frx":0E3E
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_aceptar 
         Height          =   375
         Left            =   420
         TabIndex        =   4
         ToolTipText     =   "Grabar"
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmCargo.frx":13D8
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_nuevo 
         Height          =   375
         Left            =   75
         TabIndex        =   5
         ToolTipText     =   "Nuevo"
         Top             =   165
         Width           =   360
         VariousPropertyBits=   268435483
         PicturePosition =   262148
         Size            =   "635;661"
         Picture         =   "FrmCargo.frx":1972
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.TextBox txt_descrip 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   6405
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3735
      Left            =   135
      TabIndex        =   10
      Top             =   960
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Doble Click para Editar o Eliminar"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "COD_MAESTRO3"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DESCRIP"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Abrev"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4229.858
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Abreviatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   120
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   9
      Top             =   450
      Width           =   840
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   105
      Width           =   495
   End
End
Attribute VB_Name = "FrmCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemporal  As ADODB.Recordset
Dim bNuevo      As Boolean

Private Sub btn_aceptar_Click()
    If Trim(TxtAbrev.Text & "") = "" Then
       MsgBox "Abreviatura es oblogatoria" & Chr(13) & "Se pintara en la boleta de pago", vbInformation
       Exit Sub
    End If
    If Not bNuevo Then
        If Trim(txt_cod.Text) = "" Or Len(Trim(txt_cod.Text)) = 0 Then Exit Sub
        If Trim(txt_descrip.Text) = "" Or Len(Trim(txt_descrip)) = 0 Then Exit Sub
    Else
        If Trim(txt_descrip.Text) = "" Or Len(Trim(txt_descrip)) = 0 Then Exit Sub
    End If
    
    If MsgBox("Desea Continuar?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Call Aceptar
    End If
    Call Form_Load
End Sub

Private Sub btn_editar_Click()
    Call Operacion(False)
End Sub

Private Sub btn_eliminar_Click()
    If Trim(txt_cod.Text) = "" Or Len(Trim(txt_cod.Text)) = 0 Then Exit Sub
    If MsgBox("Desea Continuar?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Call Aceptar(True)
    End If
    Call Form_Load
End Sub

Private Sub btn_nuevo_Click()
    Call Operacion(True)
    txt_cod.Text = Empty
    txt_descrip.Text = Empty
End Sub

Private Sub btn_salir_Click()
    Unload Me
End Sub

Private Sub dg_DblClick()
    If dg.ApproxCount > 0 Then
        txt_cod.Text = Trim(dg.Columns(0).Value)
        txt_descrip.Text = Trim(dg.Columns(1).Value)
        TxtAbrev.Text = Trim(dg.Columns(2).Value & "")
    End If
End Sub

Private Sub Form_Load()
    Call Inicial
    Call Trae_Info
End Sub

Private Sub Trae_Info()
    Cadena = "SP_TRAE_CARGOS_CIA '" & wcia & "'"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    Set dg.DataSource = rsTemporal
    With dg
        .Columns(0).Caption = "Codigo"
        .Columns(1).Caption = "Descripcion"
        .Columns(2).Caption = "Abrev."
        .Columns(0).Width = 1200
        .Columns(1).Width = 4200
        .Columns(2).Width = 1500
    End With
End Sub

Public Sub Aceptar(Optional ByVal Eliminar As Boolean = False)
    Cadena = "SP_MANT_CARGO '" & wcia & "', '" & txt_cod.Text & "', '" & txt_descrip.Text & "', '" & Trim(TxtAbrev.Text) & "', '" & wuser & "', " & CInt(Eliminar) & ""
    If Not EXEC_SQL(Cadena, cn) Then
        If Eliminar Then
            MsgBox "Error al eliminar el registro.", vbCritical + vbOKOnly, Me.Caption
        Else
            MsgBox "Error al Grabar el registro.", vbExclamation + vbOKOnly, Me.Caption
        End If
    Else
        If Eliminar Then
            MsgBox "Se elimino satisfactoriamente el resgisto.", vbInformation + vbOKOnly, Me.Caption
        Else
            MsgBox "Se grabó satisfactoriamente el registro.", vbInformation + vbOKOnly, Me.Caption
        End If
    End If
End Sub

Private Sub Inicial()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 5910
    Me.Width = 7920
    txt_cod.Text = Empty
    txt_descrip.Text = Empty
    txt_cod.Enabled = False
    txt_descrip.Enabled = False
    TxtAbrev.Enabled = False
    bNuevo = False
End Sub

Private Sub Operacion(ByVal bBoolean As Boolean)
    bNuevo = bBoolean
    txt_descrip.Enabled = True
    TxtAbrev.Enabled = True
    On Error Resume Next
    txt_descrip.SetFocus
End Sub

Private Sub txt_descrip_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
