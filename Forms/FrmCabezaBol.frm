VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form FrmCabezaBol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Mantenimiento de Boletas de Pago «"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8640
   Icon            =   "FrmCabezaBol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7677.217
   ScaleMode       =   0  'User
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNomSecu 
      Height          =   285
      Left            =   7065
      TabIndex        =   34
      Top             =   2100
      Width           =   1400
   End
   Begin VB.TextBox txtNomPri 
      Height          =   285
      Left            =   5655
      TabIndex        =   33
      Top             =   2115
      Width           =   1400
   End
   Begin VB.TextBox txtApeMat 
      Height          =   285
      Left            =   4200
      TabIndex        =   32
      Top             =   2130
      Width           =   1425
   End
   Begin VB.TextBox txtApePat 
      Height          =   285
      Left            =   2730
      TabIndex        =   31
      Top             =   2130
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00000000&
      Height          =   6750
      Left            =   0
      TabIndex        =   10
      Top             =   495
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "Devengadas"
         Enabled         =   0   'False
         Height          =   495
         Left            =   195
         TabIndex        =   28
         Top             =   5775
         Width           =   1095
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "QTA. CAT."
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "CTA: CTE."
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "DEP.  BCO"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "BILLETAJE"
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "BOLETA"
         BevelWidth      =   1
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Nro. registros"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label LblReg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   5160
         Width           =   1215
      End
   End
   Begin VB.TextBox Txtcodobra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1575
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   1500
      TabIndex        =   19
      Top             =   495
      Width           =   6975
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         ItemData        =   "FrmCabezaBol.frx":030A
         Left            =   4800
         List            =   "FrmCabezaBol.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   285
         Left            =   750
         TabIndex        =   2
         Top             =   705
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   94896129
         CurrentDate     =   37265
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3270
         TabIndex        =   4
         Top             =   705
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker Cmbal 
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   705
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   94896129
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Top             =   705
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   94896129
         CurrentDate     =   37267
      End
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   705
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         ItemData        =   "FrmCabezaBol.frx":030E
         Left            =   1230
         List            =   "FrmCabezaBol.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   5400
         TabIndex        =   25
         Top             =   705
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         Height          =   195
         Left            =   3720
         TabIndex        =   24
         Top             =   705
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         Height          =   195
         Left            =   2160
         TabIndex        =   23
         Top             =   705
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Boleta"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5340
      Left            =   1455
      TabIndex        =   18
      Top             =   1905
      Width           =   7095
      Begin VB.TextBox txtCodigo 
         Height          =   300
         Left            =   60
         TabIndex        =   39
         Top             =   220
         Width           =   1200
      End
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "FrmCabezaBol.frx":0312
         Height          =   4455
         Left            =   60
         TabIndex        =   6
         Top             =   555
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7858
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
         ColumnCount     =   7
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
            Caption         =   "Nombre"
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
            DataField       =   "moneda"
            Caption         =   "Moneda"
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
            DataField       =   "totneto"
            Caption         =   "Neto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Id_Boleta"
            Caption         =   "ID_Boleta"
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
         BeginProperty Column05 
            DataField       =   "fechaproceso"
            Caption         =   "Fecha"
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
         BeginProperty Column06 
            DataField       =   "id_boleta"
            Caption         =   "id_boleta"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2940.095
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adocabeza 
         Height          =   330
         Left            =   60
         Top             =   5010
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
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
      Begin VB.Label Label13 
         Caption         =   "Código"
         Height          =   450
         Left            =   60
         TabIndex        =   40
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "Seg. Nombre"
         Height          =   180
         Left            =   5655
         TabIndex        =   38
         Top             =   0
         Width           =   1365
      End
      Begin VB.Label Label11 
         Caption         =   "Pri. Nombre"
         Height          =   300
         Left            =   4230
         TabIndex        =   37
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Ap. Materno"
         Height          =   300
         Left            =   2805
         TabIndex        =   36
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label9 
         Caption         =   "Ap. Paterno"
         Height          =   285
         Left            =   1305
         TabIndex        =   35
         Top             =   0
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   60
         Width           =   5490
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6675
         TabIndex        =   20
         Top             =   120
         Width           =   1815
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
         TabIndex        =   16
         Top             =   120
         Width           =   825
      End
   End
   Begin vbalIml6.vbalImageList ilsIcons32 
      Left            =   0
      Top             =   0
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   24
      Size            =   48532
      Images          =   "FrmCabezaBol.frx":032A
      Version         =   131072
      KeyCount        =   11
      Keys            =   "FindÿSystemÿExplorerÿFavouritesÿCalendarÿNetwork NeighbourhoodÿHistoryÿInternet ExplorerÿMailÿNewsÿChannels"
   End
   Begin VB.Label Lblobra 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2595
      TabIndex        =   27
      Top             =   1590
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "FrmCabezaBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String
Dim VTipoPago As String
Dim VHorasBol As Integer

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))

If wGrupoPla = "01" And UCase(wuser) <> "SA" Then
   wciamae = Determina_Maestro("01078")
   Sql$ = "Select COD_MAESTRO2,DESCRIP from maestros_2 where status<>'*' and (cod_maestro2 in(select tipo from pla_permisos where usuario='" & wuser & "' and calculo='B'))"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Rs.RecordCount > 0 Then Rs.MoveFirst
   Do While Not Rs.EOF
      Cmbtipo.AddItem Rs!DESCRIP
      Cmbtipo.ItemData(Cmbtipo.NewIndex) = Trim(Rs!cod_maestro2)
      Rs.MoveNext
   Loop
   Rs.Close
Else
   Call fc_Descrip_Maestros2("01078", "", Cmbtipo)
End If
If Cmbtipo.ListCount = 1 Then Cmbtipo.ListIndex = 0
If wtipodoc = True Then
   Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
   Procesa_Cabeza_Boleta
Else
   wciamae = Determina_Maestro("01055")
   Sql$ = "Select * from maestros_2 where flag1='04' and status<>'*'"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Rs.RecordCount > 0 Then Rs.MoveFirst
   Do While Not Rs.EOF
      Cmbtipotrabajador.AddItem Rs!DESCRIP
      Cmbtipotrabajador.ItemData(Cmbtipotrabajador.NewIndex) = Trim(Rs!cod_maestro2)
      Rs.MoveNext
   Loop
   Rs.Close
   If Cmbtipotrabajador.ListCount >= 0 Then Cmbtipotrabajador.ListIndex = 0
End If
If wImportaUti Then Call rUbiIndCmbBox(Cmbtipo, "11", "00"): Cmbtipo.Enabled = False
End Sub

Private Sub Cmbfecha_Change()
If Month(Cmbfecha.Value) = 1 And VTipo = "02" And Cmbtipotrabajador.ListIndex >= 0 Then Command1.Enabled = True Else Command1.Enabled = False
Procesa_Cabeza_Boleta
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
Txtsemana.Text = ""
Cmbdel.Enabled = False
Cmbal.Enabled = False
Label5.Visible = False
Label6.Visible = False
Cmbdel.Visible = False
Cmbal.Visible = False
Command1.Enabled = False

If VTipo = "02" Then
   Cmbdel.Visible = False
   Cmbal.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Label4.Visible = False
   Txtsemana.Visible = False
   UpDown1.Visible = False
   Label5.Visible = True
   Label6.Visible = True
   Cmbdel.Visible = True
   Cmbal.Visible = True
   Cmbdel.Enabled = True
   Cmbal.Enabled = True
   If Month(Cmbfecha.Value) = 1 And Cmbtipotrabajador.ListIndex >= 0 Then Command1.Enabled = True Else Command1.Enabled = False
ElseIf VTipo = "03" Then
   Cmbdel.Visible = False
   Cmbal.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Label4.Visible = False
   Txtsemana.Visible = False
   UpDown1.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Cmbdel.Visible = False
   Cmbal.Visible = False
End If
Cmbtipotrabajador_Click
Procesa_Cabeza_Boleta
End Sub
Private Sub Cmbtipotrabajador_Click()
VTipotrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim wciamae As String
Dim wBeginMonth As String

VHorasBol = 0
VTipoPago = ""
wciamae = Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where cod_maestro2='" & VTipotrab & "' and status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   VHorasBol = Val(Rs!flag2)
   VTipoPago = Left(Rs!flag1, 2)
End If
If Trim(VTipoPago) = "" Then Exit Sub
If Month(Cmbfecha.Value) = 1 And VTipo = "02" Then Command1.Enabled = True Else Command1.Enabled = False
If VTipo = "01" Or VTipo = "05" Or VTipo = "11" Then
    Select Case Left(Rs!flag1, 2)
           Case Is <> "02"
                Txtsemana.Text = ""
                Txtsemana.Visible = False
                UpDown1.Visible = False
                Label4.Visible = False
                Label5.Visible = False
                Label6.Visible = False
                Cmbdel.Visible = False
                Cmbal.Visible = False
                
                Sql$ = "select iniciomes from cia where cod_cia='" & wcia & "' and status<>'*'"
                If (fAbrRst(Rs, Sql$)) Then
                   If IsNull(Rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = Rs!iniciomes
                End If
                Rs.Close
                
                If Trim(wBeginMonth) = "" Then
                    MsgBox "Ingrese el Inicio Del Mes", vbInformation, ""
                Exit Sub
                End If
                
'                Cmbfecha.Month = Month(Date)
'                Cmbfecha.Year = Year(Date)
                If Trim(wBeginMonth) <> "1" Then
                   Cmbfecha.Day = Val(wBeginMonth) - 1
                Else
                   Cmbfecha.Day = Val(fMaxDay(Cmbfecha.Month, Cmbfecha.Year))
                End If
           Case Else
                Txtsemana.Visible = True
                UpDown1.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Label6.Visible = True
                Cmbdel.Visible = True
                Cmbal.Visible = True
    End Select
End If
'Frmboleta.DgrdPagAdic.Enabled = True
If VTipotrab = "05" And VTipo = "01" Then
   Txtcodobra.Text = ""
   Lblobra.Caption = ""
   Txtcodobra.Visible = False
   Lblobra.Visible = False
  ' Frmboleta.DgrdPagAdic.Enabled = False
Else
   Txtcodobra.Text = ""
   Lblobra.Caption = ""
   Txtcodobra.Visible = False
   Lblobra.Visible = False
End If
If Rs.State = 1 Then Rs.Close
Procesa_Cabeza_Boleta
End Sub

Private Sub Command1_Click()
'Sql = "select placod from plaprovvaca where cia='" & wcia & "' and year(fechaproceso)=" & Cmbfecha.Year - 1 & " and month(fechaproceso)=12 and status<>'*'"
'If Not (fAbrRst(rs, Sql)) Then
'   MsgBox "No se encuentra la provision de Vacaciones Correspondiente", vbCritical, "Vacaciones Devengadas"
'   rs.Close
'   Exit Sub
'End If
'rs.Close
'Sql = "select placod from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Cmbfecha.Year & " and month(fechaproceso)=" & Cmbfecha.Month & " and status='D'"
'If (fAbrRst(rs, Sql)) Then
'   MsgBox "Vacaciones Devengadas ya fueron Generadas", vbInformation, "Vacaciones Devengadas"
'   rs.Close
'   Exit Sub
'End If
'rs.Close
'Call Frmboleta.Carga_Boleta("", "02", True, "", Cmbfecha.Value, VTipotrab, "04", VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, True)
    Sql = "select placod from plaprovvaca where cia='" & wcia & "' and year(fechaproceso)=" & Cmbfecha.Year - 1 & " and month(fechaproceso)=12 and status<>'*'"
    
    If Not (fAbrRst(Rs, Sql)) Then
        MsgBox "No se encuentra la provision de Vacaciones Correspondiente", vbCritical, "Vacaciones Devengadas"
        Rs.Close
        Exit Sub
    End If
    Rs.Close
    
    Sql = "select placod from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Cmbfecha.Year & " and month(fechaproceso)=" & Cmbfecha.Month & " and status='D'"
    
    If (fAbrRst(Rs, Sql)) Then
        MsgBox "Vacaciones Devengadas ya fueron Generadas", vbInformation, "Vacaciones Devengadas"
        Rs.Close
        Exit Sub
    End If
    Rs.Close
    
    Call Frmboleta.Carga_Boleta("", "02", True, "", Cmbfecha.Value, VTipotrab, "04", VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, True, "")

End Sub

Private Sub Dgrdcabeza_DblClick()
Call Frmboleta.Carga_Boleta(Trim(Dgrdcabeza.Columns(0)), VTipo, False, Txtsemana.Text, Cmbfecha.Value, VTipotrab, VTipoPago, VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, False, Dgrdcabeza.Columns(4))
End Sub

Private Sub Dgrdcabeza_HeadClick(ByVal ColIndex As Integer)
Dim xCodProd$
Select Case ColIndex
Case 0
        
        Dgrdcabeza.ClearSelCols
        
        xCodProd$ = InputBox("Digíte Cod. Trabajador a Buscar")
        If Trim(xCodProd$) = "" Then Exit Sub
        If Adocabeza.Recordset.RecordCount > 0 Then Adocabeza.Recordset.MoveFirst
        Adocabeza.Recordset.FIND "placod = '" + xCodProd$ + "'", 0, 1, 1
        If Adocabeza.Recordset.EOF Then
            MsgBox "Código No Encontrado", vbInformation
        Else
        Dgrdcabeza.SetFocus
        End If
'Case 1
'        Grd.ClearSelCols
'        Sql = "SELECT * FROM ListP ORDER BY DESCRIP"
'        datprec.RecordSource = Sql$
'        'Set datprec.Recordset = datprec.Database.OpenRecordset(SQL$, dbOpenDynaset)
'        datprec.Refresh
'
'        xCodProd$ = InputBox("Digíte Descripción a Buscar")
'        If Trim(xCodProd$) = "" Then Exit Sub
'        I = Len(xCodProd$)
'        datprec.Recordset.FindFirst "left(descrip," & I & ") = """ + xCodProd$ + """"
'        If datprec.Recordset.NoMatch Then
'            MsgBox "Descripción No Encontrado", vbInformation
'        Else
'        Grd.SetFocus
'        End If

End Select

End Sub

Private Sub Dgrdcabeza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Dgrdcabeza_DblClick
End If
End Sub

Private Sub Form_Activate()
If wtipodoc = True Then
   SSCommand1.Caption = "BOLETA"
   'Frame2.BackColor = &H80000008
   'Frame3.BackColor = &H808000
   Me.Caption = "CALCULO DE BOLETAS"
   'Label2.Visible = True
   Cmbtipo.Visible = True
Else
   SSCommand1.Caption = "ADELANTO"
   Me.Caption = "ADELANTO DE QUINCENA"
   'Frame2.BackColor = &H808000
   'Frame3.BackColor = &H80000001
   Label2.Visible = False
   Cmbtipo.Visible = False
   VTipo = "01"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    'If Not TypeOf Screen.ActiveControl Is DataGrid Then
        Sendkeys "{TAB}"
    'Else
        'Dgrdcabeza_DblClick
    'End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 8730
Me.Height = 7710

Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")

Cmbfecha.Month = Month(Date)
Cmbfecha.Year = Year(Date)
Cmbfecha.Day = Day(Date)


Cmbdel.Year = Year(Date)
Cmbdel.Month = Month(Date)
Cmbdel.Day = Day(Date)

Cmbal.Year = Year(Date)
Cmbal.Month = Month(Date)
Cmbal.Day = Day(Date)


'Dim barX As cListBar
'Dim itmX As cListBarItem
'Dim i As Long
'
'   With vbalListBar1
'
'      .ImageList(evlbLargeIcon) = ilsIcons32
'
'      Set barX = .Bars.Add("OPCIONES", , " ")
'
'      Set itmX = barX.Items.Add("boleta", , "Boleta", 9)
'
'      Set itmX = barX.Items.Add("billetaje", , "Billetaje", 2)
'
'      Set itmX = barX.Items.Add("depbco", , "Dep. Bco", 4)
'
'      Set itmX = barX.Items.Add("ctacte", , "Cta. Cte.", 9)
'
'      Set itmX = barX.Items.Add("qtacta", , "Qta. Cat.", 2)
'
'      Set itmX = barX.Items.Add("devengadas", , "Devengadas", 4)
'   End With

End Sub

Private Sub SSCommand1_Click()
Txtsemana.Text = Format(Txtsemana.Text, "00")

Sql$ = "select uit from plauit where cia='" & wcia & "' and ano=" & Cmbfecha.Year & " and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   If Rs(0) <= 0 Then MsgBox "Debe registrar UIT", vbInformation: Exit Sub
Else
   MsgBox "Debe registrar UIT", vbInformation
   Exit Sub
End If
If Rs.State = 1 Then Rs.Close

If wtipodoc = True Then

    'mgirao solo practicantes van a entrar al if
    If VTipotrab <> "05" Then
        If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Indicar Tipo de Boleta", vbCritical, TitMsg: Cmbtipo.SetFocus: Exit Sub
        If Cmbtipotrabajador.ListIndex < 0 Then MsgBox "Debe Indicar Tipo de Trabajador", vbCritical, TitMsg: Cmbtipotrabajador.SetFocus: Exit Sub
        If VTipoPago = "02" And (VTipo = "01" Or VTipo = "11") And Txtsemana.Text = "" Then MsgBox "Debe Indicar Semana", vbCritical, TitMsg: Txtsemana.SetFocus: Exit Sub
        If Txtcodobra.Visible = True And Txtcodobra.Text = "" Then MsgBox "Debe Indicar Obra", vbCritical, TitMsg: Txtsemana.SetFocus: Exit Sub
    End If
    
    Call Frmboleta.Carga_Boleta("", VTipo, True, Txtsemana.Text, Cmbfecha.Value, VTipotrab, VTipoPago, VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, False, "")
    Frmboleta.Caption = Frmboleta.Caption & " " & Cmbtipo.Text & " - " & Cmbtipotrabajador.Text
    If VTipo = "02" Then
        Frmboleta.Txtvaca.Text = Cmbfecha.Value
        Frmboleta.Txtvacai.Text = Cmbdel.Value
        Frmboleta.Txtvacaf.Text = Cmbal.Value
    End If
Else
    
    Call Frmboleta.Carga_Boleta("", "01", True, "", Cmbfecha.Value, VTipotrab, "04", VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, False, "")
End If

End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtApeMat_Change()
Procesa_Cabeza_Boleta
End Sub

Private Sub TxtApePat_Change()
Procesa_Cabeza_Boleta
End Sub
Private Sub txtCodigo_Change()
Procesa_Cabeza_Boleta
End Sub

Private Sub Txtcodobra_GotFocus()
NameForm = "FrmCabezaBol"
wbus = "OB"
End Sub

Private Sub Txtcodobra_KeyPress(KeyAscii As Integer)
Txtcodobra.Text = Txtcodobra.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Txtcodobra.Text = Format(Txtcodobra.Text, "00000000"): Cmbtipo.SetFocus
End Sub

Private Sub Txtcodobra_LostFocus()
wbus = ""
If Txtcodobra.Text <> "" Then
   Sql$ = "select cod_obra,descrip,status from plaobras where cod_cia='" & wcia & "' and cod_obra='" & Txtcodobra.Text & "'"
   If (fAbrRst(Rs, Sql$)) Then
      If Rs!status = "*" Then
         MsgBox "Obra Eliminada", vbInformation, "Registro de Personal"
         Lblobra.Caption = ""
         Txtcodobra.SetFocus
      Else
         Lblobra.Caption = Trim(Rs!DESCRIP)
      End If
   Else
     MsgBox "Codigo de Obra no Registrada", vbInformation, "Registro de Personal"
     Lblobra.Caption = ""
     Txtcodobra.SetFocus
   End If
End If
End Sub

Private Sub txtNomPri_Change()
Procesa_Cabeza_Boleta
End Sub

Private Sub txtNomSecu_Change()
Procesa_Cabeza_Boleta
End Sub

Private Sub Txtsemana_Change()
 Procesa_Cabeza_Boleta
End Sub

Private Sub Txtsemana_KeyPress(KeyAscii As Integer)
Txtsemana.Text = Txtsemana.Text + fc_ValNumeros(KeyAscii)
End Sub

Private Sub UpDown1_DownClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "00"
If Txtsemana.Text > 0 Then Txtsemana.Text = Format(Val(Txtsemana.Text - 1), "00")
End Sub

Private Sub UpDown1_UpClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "00"
Txtsemana.Text = Format(Val(Txtsemana.Text + 1), "00")


End Sub
Public Sub Procesa_Cabeza_Boleta()
Dim mano As Integer
Dim mmes As Integer
On Error GoTo CORRIGE
Me.LblReg.Caption = "0"

If Trim(Txtsemana.Text) <> "" Then
   Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(Cmbfecha.Year, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   
   Set Rs = cn.Execute(Sql$, 64)
   
   If Rs.RecordCount > 0 Then
      Cmbdel.Value = Format(Rs!fechai, "dd/mm/yyyy")
      Cmbal.Value = Format(Rs!fechaf, "dd/mm/yyyy")
      Cmbfecha.Value = Format(Rs!fechaf, "dd/mm/yyyy")
      
    End If
   
   If Rs.State = 1 Then Rs.Close
End If

Dgrdcabeza.Enabled = False

If Trim(VTipoPago) = "" Or IsNull(VTipoPago) Then Exit Sub

   mano = Val(Mid(Cmbfecha.Value, 7, 4))
   mmes = Val(Mid(Cmbfecha.Value, 4, 2))
   
   If wtipodoc = True Then
      Select Case VTipoPago
        Case Is = "02"
             Dim strMes As String
             strMes = ""
             'If Txtsemana.Visible = False Then
             '   If MsgBox("Desea Ver Todas las Boletas del Mes?", vbInformation + vbYesNo + vbDefaultButton2, "Planilla") = vbNo Then
             '      StrMes = " and day(fechaproceso)=" & Format(Cmbfecha.Day, "00") & ""
             '   End If
             'End If
             
                Sql$ = nombre()
                Sql$ = Sql$ & "a.id_boleta,a.placod,a.totneto,b.pagomoneda as moneda,a.fechaproceso " _
                & " from plahistorico a,planillas b " _
                & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' " _
                & " and ap_pat like '" & Trim(txtApePat.Text) + "%" & "'  and b.placod like '" & Trim(txtCodigo.Text) + "%" & "'  and ap_mat like '" & Trim(txtApeMat.Text) + "%" & "' and nom_1 like '" & Trim(txtNomPri.Text) + "%" & "'  and nom_2 like '" & Trim(txtNomSecu.Text) + "%" & "'" _
                & " and a.semana='" & Txtsemana.Text & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " "
                Sql$ = Sql$ & strMes
                Sql$ = Sql$ & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*'" _
                & " order by a.placod"
                
        Case Is = "04"
             Sql$ = nombre()
             Sql$ = Sql$ & "a.id_boleta,a.placod,a.totneto,b.pagomoneda as moneda,a.fechaproceso " _
             & " from plahistorico a,planillas b " _
             & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' " _
             & " and ap_pat like '" & Trim(txtApePat.Text) + "%" & "'  and b.placod like '" & Trim(txtCodigo.Text) + "%" & "'  and ap_mat like '" & Trim(txtApeMat.Text) + "%" & "' and nom_1 like '" & Trim(txtNomPri.Text) + "%" & "'  and nom_2 like '" & Trim(txtNomSecu.Text) + "%" & "'" _
             & " and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
             & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' " _
             & " order by a.placod"
      
      End Select
   
   Else
      Sql$ = nombre()
      Sql$ = Sql$ & "a.placod,a.totneto,b.pagomoneda as moneda " _
      & "from plaquincena a,planillas b " _
      & "where a.cia='" & wcia & "' and b.tipotrabajador='" & VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
      & " and ap_pat like '" & Trim(txtApePat.Text) + "%" & "'  and b.placod like '" & Trim(txtCodigo.Text) + "%" & "'  and ap_mat like '" & Trim(txtApeMat.Text) + "%" & "' and nom_1 like '" & Trim(txtNomPri.Text) + "%" & "'  and nom_2 like '" & Trim(txtNomSecu.Text) + "%" & "'" _
      & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' " & _
      " order by a.placod"
   End If

 cn.CursorLocation = adUseClient
 Set Adocabeza.Recordset = cn.Execute(Sql$, 64)
 
 If Adocabeza.Recordset.RecordCount > 0 Then
    Adocabeza.Recordset.MoveFirst
    Dgrdcabeza.Enabled = True
 Else
    Dgrdcabeza.Enabled = False
 End If
 Dgrdcabeza.Refresh
 Me.LblReg.Caption = Adocabeza.Recordset.RecordCount
 
 Screen.MousePointer = vbDefault
 Exit Sub
 
CORRIGE:
 MsgBox "Error :" & Err.Description, vbCritical, Me.Caption
 
End Sub

Private Sub vbalListBar1_ItemClick(Item As vbalLbar6.cListBarItem, Bar As vbalLbar6.cListBar)
Select Case Item.Key
    
Case "boleta"
    If wtipodoc = True Then
    If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Indicar Tipo de Boleta", vbCritical, TitMsg: Cmbtipo.SetFocus: Exit Sub
    If Cmbtipotrabajador.ListIndex < 0 Then MsgBox "Debe Indicar Tipo de Trabajador", vbCritical, TitMsg: Cmbtipotrabajador.SetFocus: Exit Sub
    If VTipoPago = "02" And VTipo = "01" And Txtsemana.Text = "" Then MsgBox "Debe Indicar Semana", vbCritical, TitMsg: Txtsemana.SetFocus: Exit Sub
    If Txtcodobra.Visible = True And Txtcodobra.Text = "" Then MsgBox "Debe Indicar Obra", vbCritical, TitMsg: Txtsemana.SetFocus: Exit Sub
        Call Frmboleta.Carga_Boleta("", VTipo, True, Txtsemana.Text, Cmbfecha.Value, VTipotrab, VTipoPago, VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, False, "")
    Else
        Call Frmboleta.Carga_Boleta("", "01", True, "", Cmbfecha.Value, VTipotrab, "04", VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, False, "")
    End If

Case "billetaje"

Case "depbco"

Case "ctacte"

Case "qtacta"

Case "devengadas"
    Sql = "select placod from plaprovvaca where cia='" & wcia & "' and year(fechaproceso)=" & Cmbfecha.Year - 1 & " and month(fechaproceso)=12 and status<>'*'"
    
    If Not (fAbrRst(Rs, Sql)) Then
        MsgBox "No se encuentra la provision de Vacaciones Correspondiente", vbCritical, "Vacaciones Devengadas"
        Rs.Close
        Exit Sub
    End If
    Rs.Close
    
    Sql = "select placod from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Cmbfecha.Year & " and month(fechaproceso)=" & Cmbfecha.Month & " and status='D'"
    
    If (fAbrRst(Rs, Sql)) Then
        MsgBox "Vacaciones Devengadas ya fueron Generadas", vbInformation, "Vacaciones Devengadas"
        Rs.Close
        Exit Sub
    End If
    Rs.Close
    
    Call Frmboleta.Carga_Boleta("", "02", True, "", Cmbfecha.Value, VTipotrab, "04", VHorasBol, Cmbdel.Value, Cmbal.Value, Txtcodobra.Text, True, "")
    
End Select

End Sub
