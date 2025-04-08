VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmgrdpla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Listado de Personal «"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   Icon            =   "Frmgrdpla.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid Dgrdplanilla 
      Bindings        =   "Frmgrdpla.frx":030A
      Height          =   4410
      Left            =   45
      TabIndex        =   5
      Top             =   780
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "placod"
         Caption         =   "Codigo"
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
         DataField       =   "nombre"
         Caption         =   "Nombres"
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
         DataField       =   "descrip"
         Caption         =   "Tipo"
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
      BeginProperty Column03 
         DataField       =   "codauxinterno"
         Caption         =   "codinterno"
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
      BeginProperty Column04 
         DataField       =   "tipotrabajador"
         Caption         =   "TipoTrab"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5160.189
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8415
      Begin VB.TextBox Txtnom2 
         Height          =   285
         Left            =   6540
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Txtnom1 
         Height          =   285
         Left            =   4860
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Txtape2 
         Height          =   285
         Left            =   3180
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Txtape1 
         Height          =   285
         Left            =   1500
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Txtcod 
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbltipotra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2do Nombre"
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
         Left            =   6540
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1er Nombre"
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
         Left            =   4860
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Materno"
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
         Left            =   3180
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Paterno"
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
         Left            =   1500
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adopla 
      Height          =   375
      Left            =   840
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Maestro de Personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   8295
   End
End
Attribute VB_Name = "Frmgrdpla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dgrdplanilla_DblClick()

Select Case MDIplared.ActiveForm.Name
        
        Case Is = "FrmCertretecionApf"
            FrmCertretecionApf.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
            FrmCertretecionApf.LblNomTrab = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "aFrmDH"
            aFrmDH.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
            aFrmDH.LblNomTrab.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "FrmConDerechoHab"
            FrmConDerechoHab.TxtNro.Text = Trim(Dgrdplanilla.Columns(0).Text)
            FrmConDerechoHab.LblNomTrab.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "FrmDiasNoTrab"
            FrmDiasNoTrab.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
            FrmDiasNoTrab.LblNomTrab.Caption = Trim(Dgrdplanilla.Columns(1).Text)
            FrmDiasNoTrab.Carga_Susidio
        Case Is = "FrmDiasSub"
            FrmDiasSub.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
            FrmDiasSub.LblNomTrab.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "FrmComisionVendedor"
            FrmComisionVendedor.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
            FrmComisionVendedor.LblNomTrab.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "Frmplacte"
            Frmplacte.Txtcodpla.Text = Trim(Dgrdplanilla.Columns(0).Text)
            Frmplacte.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "Frmgrdctacte"
           'Frmgrdctacte.Txtcod.Text = Trim(Dgrdplanilla.Columns(0).Text)
           'Frmgrdctacte.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "Frmsubsidios"
            Frmsubsidios.Txtcodpla.Text = Trim(Dgrdplanilla.Columns(0).Text)
            Call Frmsubsidios.Txtcodpla_KeyPress(13)
        
        Case Is = "Frmboleta"
           Frmboleta.Txtcodpla.Text = Trim(Dgrdplanilla.Columns(0).Text)
           Frmboleta.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
           Frmboleta.Lblcodaux.Caption = Trim(Dgrdplanilla.Columns(3).Text)
           Frmboleta.Txtcodpla_KeyPress (13)
           
        Case Is = "frmcontratos"
           frmcontratos.TxtcodSuplencia.Text = Trim(Dgrdplanilla.Columns(0).Text)
           frmcontratos.Label14.Caption = Trim(Dgrdplanilla.Columns(1).Text)
           'frmcontratos.TxtcodSuplencia_KeyPress (13)
        Case Is = "FrmPromedios"
           FrmPromedios.TxtCodTrabajador.Text = Trim(Dgrdplanilla.Columns(0).Text)
           'FrmPromedios.TxtCodTrabajador_KeyPress (13)
        Case "Frmprovision"
           Frmprovision.Txtcodpla.Text = Trim(Dgrdplanilla.Columns(0).Text)
           Frmprovision.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
           Frmprovision.LblTipoTrab = Trim(Dgrdplanilla.Columns(4).Text)
        Case Is = "Frmtareo"
           Frmtareo.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
           Frmtareo.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "Frmgrdtareo"
           Frmgrdtareo.TxtCodTrab.Text = Trim(Dgrdplanilla.Columns(0).Text)
           Frmgrdtareo.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "Frmdetalle"
           Frmdetalle.TxtCodigo.Text = Trim(Dgrdplanilla.Columns(0).Text)
           Frmdetalle.Lblnombre.Caption = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "FrmCtsFalta"
            FrmCtsFalta.txtplacod.Text = Trim(Dgrdplanilla.Columns(0).Text)
            FrmCtsFalta.txtpersonal.Text = Trim(Dgrdplanilla.Columns(1).Text)
        Case Is = "FrmUtilidades"
            If FrmUtilidades.FrameCtaCte.Visible = True Then
               FrmUtilidades.TxtCodCtaCte.Text = Trim(Dgrdplanilla.Columns(0).Text)
               FrmUtilidades.LblTrabajadorCtaCte.Caption = Trim(Dgrdplanilla.Columns(1).Text)
            End If
        Case Is = "FrmTrabRemOtrasEmp"
            If FrmTrabRemOtrasEmp.SSpanelNewTrab.Visible = True And FrmTrabRemOtrasEmp.BtnAceptar.Visible = True Then
               FrmTrabRemOtrasEmp.TxtCodNewTrab = Trim(Dgrdplanilla.Columns(0).Text)
               FrmTrabRemOtrasEmp.LblNombreNewTrab = Trim(Dgrdplanilla.Columns(1).Text)
            End If
         Case Is = "FrmAgregaSegVida"
             FrmAgregaSegVida.Agrega_Persona (Trim(Dgrdplanilla.Columns(0).Text))
        End Select
        Unload Me
End Sub

Private Sub Dgrdplanilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Dgrdplanilla.VisibleRows > 0 Then Dgrdplanilla_DblClick
End Sub

Private Sub Form_Activate()
    Me.Width = 8385
    Me.Height = 5610
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Sql$ = SP_PLANILLAS_NOMBRE(wcia, Trim(Txtape1.Text), Trim(Txtape2.Text), Trim(Txtnom1.Text), Trim(Txtnom2.Text))
'Debug.Print SQL$
cn.CursorLocation = adUseClient

Set Adopla.Recordset = cn.Execute(Sql, 64)
If Adopla.Recordset.EOF Then MsgBox ("No Existe Personal en Planilla de la Compañia")
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIplared.ActiveForm.Enabled = True
End Sub

Private Sub Txtape1_Change()
   Screen.MousePointer = vbHourglass
   Sql$ = SP_PLANILLAS_NOMBRE(wcia, Trim(Txtape1.Text), Trim(Txtape2.Text), Trim(Txtnom1.Text), Trim(Txtnom2.Text))
   Set Adopla.Recordset = cn.Execute(Sql$, 64)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Txtape1_GotFocus()
    Txtcod.Text = ""
End Sub
Private Sub Txtape1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn And Dgrdplanilla.VisibleRows > 0 Then Dgrdplanilla.SetFocus
End Sub

Private Sub Txtape2_Change()
   Screen.MousePointer = vbHourglass
   Sql$ = SP_PLANILLAS_NOMBRE(wcia, Trim(Txtape1.Text), Trim(Txtape2.Text), Trim(Txtnom1.Text), Trim(Txtnom2.Text))
   Set Adopla.Recordset = cn.Execute(Sql$, 64)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Txtape2_GotFocus()
    Txtcod.Text = ""
End Sub

Private Sub txtcod_Change()
    Dim mcod As String
    mcod = Trim(Txtcod.Text)
    Txtcod.Text = mcod
    Sql$ = SP_PLANILLAS_CODIGO(wcia, mcod)
    Set Adopla.Recordset = cn.Execute(Sql$, 64)
    If Adopla.Recordset.RecordCount = 0 Then
        Txtcod.Text = mcod
        Txtcod.SelStart = 0
        Txtcod.SelLength = Len(Txtcod.Text)
        Dgrdplanilla.Refresh
        MsgBox "El Codigo de Personal No existe", vbExclamation, "Verifique"
    End If
    Txtcod.SetFocus
End Sub

Private Sub Txtcod_GotFocus()
    Txtape1.Text = ""
    Txtape2.Text = ""
    Txtnom1.Text = ""
    Txtnom1.Text = ""
End Sub

Private Sub Txtcod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn And Dgrdplanilla.VisibleRows > 0 Then Dgrdplanilla.SetFocus
End Sub

Private Sub Txtcod_KeyPress(KeyAscii As Integer)
'Dim mcod As String
'If KeyAscii = 13 And Txtcod.Text <> "" Then
'mcod = Trim(Txtcod.Text)
'Txtcod.Text = mcod
'Screen.MousePointer = vbHourglass
'    Sql$ = SP_PLANILLAS_CODIGO(wcia, mcod)
'    Set Adopla.Recordset = cn.Execute(Sql$, 64)
'    If Adopla.Recordset.RecordCount = 0 Then
'       Txtcod.Text = mcod
'       Dgrdplanilla.Refresh
'       MsgBox "El Codigo de Personal No existe", vbExclamation, "Verifique"
'    End If
'    Txtcod.SelStart = 0
'    Txtcod.SelLength = Len(Txtcod.Text)
'    Txtcod.SetFocus
'    Screen.MousePointer = vbDefault
'    Exit Sub
'End If
End Sub

Private Sub Txtnom1_Change()
   Screen.MousePointer = vbHourglass
   Sql$ = SP_PLANILLAS_NOMBRE(wcia, Trim(Txtape1.Text), Trim(Txtape2.Text), Trim(Txtnom1.Text), Trim(Txtnom2.Text))
   Set Adopla.Recordset = cn.Execute(Sql$, 64)
   Screen.MousePointer = vbDefault

End Sub

Private Sub Txtnom1_GotFocus()
Txtcod.Text = ""
End Sub

Private Sub Txtnom1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn And Dgrdplanilla.VisibleRows > 0 Then Dgrdplanilla.SetFocus
End Sub

Private Sub Txtnom2_Change()
   Screen.MousePointer = vbHourglass
   Sql$ = SP_PLANILLAS_NOMBRE(wcia, Trim(Txtape1.Text), Trim(Txtape2.Text), Trim(Txtnom1.Text), Trim(Txtnom2.Text))
   Set Adopla.Recordset = cn.Execute(Sql$, 64)
   Screen.MousePointer = vbDefault

End Sub

Private Sub Txtnom2_GotFocus()
Txtcod.Text = ""
End Sub

Private Sub Txtnom2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn And Dgrdplanilla.VisibleRows > 0 Then Dgrdplanilla.SetFocus
End Sub
