VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmprintcrystal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresion de Boletas."
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox CRViewer91 
      Height          =   6720
      Left            =   7290
      ScaleHeight     =   6660
      ScaleWidth      =   5355
      TabIndex        =   29
      Top             =   45
      Width           =   5415
   End
   Begin VB.TextBox Txtcodobra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   6975
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   51838977
         CurrentDate     =   37265
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   660
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker Cmbal 
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51838977
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51838977
         CurrentDate     =   37267
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   3600
         TabIndex        =   25
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   5400
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         Height          =   195
         Left            =   3720
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Boleta"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4695
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   7215
      Begin VB.Frame FramePrint 
         Height          =   1695
         Left            =   2520
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton Opcindividual 
            Caption         =   "Individual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   240
            Width           =   1320
         End
         Begin VB.OptionButton Opctotal 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   600
            Width           =   945
         End
         Begin VB.OptionButton Opcrango 
            Caption         =   "A partir de Seleccion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   960
            Width           =   2040
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Imprimir Boletas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   20
            TabIndex        =   6
            Top             =   1300
            Width           =   2130
         End
      End
      Begin MSAdodcLib.Adodc Adocabeza 
         Height          =   375
         Left            =   1200
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   945
         TabIndex        =   10
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "frmprintcrystal.frx":0000
         Height          =   4455
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7858
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
            DataField       =   "fechaproceso"
            Caption         =   "fechaproceso"
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
            BeginProperty Column00 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3690.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   4455
      End
      Begin VB.Label Lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Lblobra 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   5880
   End
   Begin VB.Label LblTipDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmprintcrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String
Dim VTipoPago As String
Dim VHorasBol As Integer
Dim rsboleta As New Recordset
Dim mlinea As Integer
Public wmBolQuin As String
Dim FLAG As Boolean
Dim CRREPORTE As New CRAXDRT.Report
Dim CNA As New CRAXDRT.Application

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01078", "", Cmbtipo)

If wmBolQuin = "B" = True Then
   Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
   Procesa_Seteo_Boleta
Else
   wciamae = Determina_Maestro("01055")
   Sql$ = "Select * from maestros_2 where flag1='04' and status<>'*'"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then rs.MoveFirst
   Do While Not rs.EOF
      Cmbtipotrabajador.AddItem rs!DESCRIP
      Cmbtipotrabajador.ItemData(Cmbtipotrabajador.NewIndex) = Trim(rs!COD_MAESTRO2)
      rs.MoveNext
   Loop
   rs.Close
   If Cmbtipotrabajador.ListCount >= 0 Then Cmbtipotrabajador.ListIndex = 0
End If
Crea_Rs
End Sub

Private Sub Cmbfecha_Change()
Procesa_Seteo_Boleta
End Sub

Private Sub CmbTipo_Click()

VTipo = Funciones.fc_CodigoComboBox(Cmbtipo, 2)
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
   
Else
   Cmbdel.Visible = True
   Cmbal.Visible = True
   Label5.Visible = True
   Label6.Visible = True
End If
Cmbtipotrabajador_Click
Procesa_Seteo_Boleta
End Sub

Private Sub Cmbtipotrabajador_Click()

VTipotrab = Funciones.fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim wciamae As String

wciamae = Funciones.Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where cod_maestro2='" & _
VTipotrab & "' and status<>'*'"

Sql$ = Sql$ & wciamae

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then VHorasBol = Val(rs!flag2): VTipoPago = Left(rs!flag1, 2)
If VTipo = "01" Then
   If rs.RecordCount > 0 Then
      Select Case Left(rs!flag1, 2)
             Case Is <> "02"
                  Txtsemana.Text = ""
                  Txtsemana.Visible = False
                  UpDown1.Visible = False
                  Label4.Visible = False
                  Label5.Visible = False
                  Label6.Visible = False
                  Cmbdel.Visible = False
                  Cmbal.Visible = False
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
End If
If VTipotrab = "05" And VTipo = "01" Then
   Txtcodobra.Text = ""
   Lblobra.Caption = ""
   Txtcodobra.Visible = True
   Lblobra.Visible = True
Else
   Txtcodobra.Text = ""
   Lblobra.Caption = ""
   Txtcodobra.Visible = False
   Lblobra.Visible = False
End If
If rs.State = 1 Then rs.Close
Procesa_Seteo_Boleta
End Sub
Private Sub Command1_Click()
If wmBolQuin = "B" Then
   'Barra.Caption = "Generando Boletas de Pago"
   Print_Bol
Else
   'Barra.Caption = "Generando Recibos de Quincena"
   Print_Quin
End If
End Sub

Private Sub Form_Activate()
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
If wmBolQuin = "B" Then
   Me.Caption = "IMPRESION DE BOLETAS"
   Frame3.BackColor = &H80000001
   Label2.Visible = True
   Cmbtipo.Visible = True
Else
   Me.Caption = "IMPRESION DE QUINCENAS"
   Frame3.BackColor = &H80000008
   Label2.Visible = False
   Cmbtipo.Visible = False
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 And FramePrint.Visible = True Then
           FramePrint.Visible = False
        End If
End Sub

Private Sub Form_Load()
FLAG = False
Me.Top = 0
Me.Left = 0
Me.Width = 7320
Me.Height = 7185
Call Funciones.rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call Funciones.rUbiIndCmbBox(Cmbcia, wcia, "00")

Cmbfecha.Year = Year(Date)
Cmbfecha.Month = Month(Date)
Cmbfecha.Day = Day(Date)

Cmbdel.Year = Year(Date)
Cmbdel.Month = Month(Date)
Cmbdel.Day = Day(Date)

Cmbal.Year = Year(Date)
Cmbal.Month = Month(Date)
Cmbal.Day = Day(Date)

Me.KeyPreview = True
End Sub

Private Sub Txtsemana_Change()
Procesa_Seteo_Boleta
End Sub
Public Sub Procesa_Seteo_Boleta()
Dim mano As Integer
Dim mmes As Integer
On Error GoTo CORRIGE

If Trim(Txtsemana.Text) <> "" Then
Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(Cmbfecha.Year, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
 
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then
   Cmbdel.Value = Format(rs!fechai, "dd/mm/yyyy")
   Cmbal.Value = Format(rs!fechaf, "dd/mm/yyyy")
End If
If rs.State = 1 Then rs.Close
End If
If wmBolQuin = "B" Then
   If VTipoPago = "" Or IsNull(VTipoPago) Then Exit Sub
End If
mano = Val(Mid(Cmbfecha.Value, 7, 4))
mmes = Val(Mid(Cmbfecha.Value, 4, 2))
If wmBolQuin = "B" Then
   Select Case VTipoPago
          Case Is = "02"
               Sql$ = nombre()
               Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso " _
                  & "from plahistorico a,planillas b " _
                  & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' and a.semana='" & Txtsemana.Text & "' " _
                  & "and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
          Case Is = "04"
                  Sql$ = nombre()
                  Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso " _
                  & "from plahistorico a,planillas b " _
                  & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
                  & "and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
   End Select
Else
   Sql$ = nombre()
   Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso " & _
   "from plaquincena a,planillas b " _
   & "where a.cia='" & wcia & "' and b.tipotrabajador='" & VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
   & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
End If
cn.CursorLocation = adUseClient
Set Adocabeza.Recordset = cn.Execute(Sql$, 64)
FLAG = True
If Adocabeza.Recordset.RecordCount > 0 Then Adocabeza.Recordset.MoveFirst: Cmbfecha.Value = Format(Adocabeza.Recordset!FechaProceso, "dd/mm/yyyy")
Dgrdcabeza.Refresh
Screen.MousePointer = vbDefault
Exit Sub
CORRIGE:
MsgBox "Error : " & ERR.Description, vbCritical, Me.Caption
End Sub


Private Sub UpDown1_DownClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "0"
If Txtsemana.Text > 0 Then Txtsemana = Txtsemana - 1
End Sub

Private Sub UpDown1_UpClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "0"
Txtsemana = Txtsemana + 1
End Sub
Public Sub Imprime_Boletas()

If Not FLAG Then Exit Sub
If Adocabeza.Recordset.RecordCount <= 0 Then Exit Sub
If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
If wmBolQuin = "B" Then Command1.Caption = "Imprimir Boletas" Else Command1.Caption = "Imprimir Quincenas"
FramePrint.Visible = True
End Sub
Private Sub Crea_Rs()
    If rsboleta.State = 1 Then rsboleta.Close
    rsboleta.Fields.Append "texto", adChar, 65, adFldIsNullable
    rsboleta.Open
End Sub
Private Function No_Seteados(tipo As String, seteo As String, status As String) As Currency
Dim mHasta As Integer
Dim mFields As Integer
Dim mCadSet As String
Dim I As Integer
Dim J As Integer
Dim mFound As Boolean
Dim mLen As Integer
Dim mcadadd As String
Dim rsnoset As ADODB.Recordset
Dim rsnorem As ADODB.Recordset
Dim mbasico As Currency

No_Seteados = 0
mLen = 0
If Trim(seteo) <> "" Then mLen = Len(Trim(seteo))

Select Case tipo
       Case Is = "IN": mHasta = 50: mFields = 44
       Case Is = "AP": mHasta = 20: mFields = 114
       Case Is = "DE": mHasta = 20: mFields = 94
End Select
mCadSet = ""
For I = 1 To mHasta
    If rs(I + mFields) <> 0 Then
       mCount = 0
       mFound = False
       If mLen > 0 Then
          For J = 1 To mLen - 1 Step 2
              If Mid(seteo, J, 2) = Format(I, "00") Then mFound = True: Exit For
          Next
          If mFound = False Then mCadSet = mCadSet & "'" & Format(I, "00") & "',"
       Else
          mCadSet = mCadSet & "'" & Format(I, "00") & "',"
       End If
    End If
Next
If Trim(mCadSet) = "" Then Exit Function

mCadSet = Mid(mCadSet, 1, Len(Trim(mCadSet)) - 1)
mCadSet = "in(" & mCadSet & ")"
    
If status = "R" Then
   Sql = "SELECT distinct(codinterno),descripcion FROM PLACONSTANTE  c, plaafectos a " _
       & "WHERE c.cia='" & wcia & "' and c.CODINTERNO " & mCadSet & " AND c.TIPOMOVIMIENTO='02'  and c.status<>'*' " _
       & "and a.cia=c.cia and c.codinterno=a.cod_remu and a.status<>'*'"
   If (fAbrRst(rsnoset, Sql)) Then rsnoset.MoveFirst
   Do While Not rsnoset.EOF
      mbasico = 0
      mnumh = Remun_Horas(rsnoset!codinterno)
      If mnumh <> 0 Then mbasico = rs(14 + mnumh)
      If mbasico <> 0 Then
         mcadadd = lentexto(10, Left(rsnoset!Descripcion, 10)) & fCadNum(mbasico, "##0.00")
      Else
         mcadadd = lentexto(16, Left(rsnoset!Descripcion, 16))
      End If
      mbasico = rs(mFields + Val(rsnoset!codinterno))
      mcadadd = mcadadd & Space(1) & fCadNum(mbasico, "###,##0.00")
      rsboleta.AddNew
      rsboleta!texto = mcadadd
      No_Seteados = No_Seteados + mbasico
      rsnoset.MoveNext
   Loop
   rsnoset.Close
ElseIf status = "N" Then
   Sql = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno " & mCadSet & " and status<>'*'"
   If (fAbrRst(rsnoset, Sql)) Then rsnoset.MoveFirst
   Do While Not rsnoset.EOF
      mbasico = 0
      Sql = "select * from plaafectos where cia='" & wcia & "' and cod_remu ='" & rsnoset!codinterno & "' and status<>'*'"
      If Not (fAbrRst(rsnorem, Sql)) Then
         mbasico = rs(mFields + Val(rsnoset!codinterno))
         mcadadd = lentexto(16, Left(rsnoset!Descripcion, 16))
         mcadadd = mcadadd & Space(1) & fCadNum(mbasico, "###,##0.00")
         rsboleta.AddNew
         rsboleta!texto = mcadadd
         No_Seteados = No_Seteados + mbasico
      End If
      rsnorem.Close
      rsnoset.MoveNext
   Loop
   rsnoset.Close
Else
   Sql = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno " & mCadSet & " and status<>'*'"
   If (fAbrRst(rsnoset, Sql)) Then rsnoset.MoveFirst
   Do While Not rsnoset.EOF
      mbasico = rs(mFields + Val(rsnoset!codinterno))
      If tipo = "AP" Then
         mcad = lentexto(12, Left(rsnoset!Descripcion, 12)) & Space(1) & fCadNum(mbasico, "##,##0.00")
      Else
         mcad = lentexto(12, Left(rsnoset!Descripcion, 12)) & Space(13) & fCadNum(mbasico, "##,##0.00")
      End If
      rsboleta.AbsolutePosition = mlinea
      rsboleta!texto = Left(rsboleta!texto, 27) & Space(4) & mcad
      mlinea = mlinea + 1
      rsnoset.MoveNext
   Loop
   rsnoset.Close
End If
End Function
Private Sub Print_Bol()
Dim mano As Integer
Dim mmes As Integer
Dim mperiodo As String, cod As String
Dim mnombre As String, fecha As String
Dim mbasico As Currency
Dim mcargo As String
Dim mcad As String
Dim mafp As String
Dim mconcep As String
Dim rs2 As ADODB.Recordset
Dim rsremu As ADODB.Recordset
Dim wciamae As String
Dim mesnombre As String
Dim mcianom As String
Dim mciaregpat As String
Dim mciaruc As String
Dim mciadir As String
Dim mciadist As String
Dim mnumh As Integer
Dim totremu As Currency
Dim mtexto As String
Dim mCadBlanc As String
Dim mCadSeteo As String
Dim RX As New ADODB.Recordset
Dim FACTOR_HORAS As Variant
Dim NIC As Variant
Dim tempo As Double, tempo1 As Double

On Error GoTo FUNKA
mtexto = ""

Set CRREPORTE = CNA.OpenReport("C:\rptprueba.rpt", 1)
Call Limpiar

If VTipoPago = "" Or IsNull(VTipoPago) Then Exit Sub
'Sql$ = "select a.*,dist from cia a,ubigeos b where cod_cia='" & wcia & "' and a.cod_ubi=b.cod_ubi"
Sql$ = " select c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dp, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais from cia c, sunat_ubigeo u " & _
      "  WHERE cod_cia='" & wcia & "' AND u.id_ubigeo=*c.cod_ubi"

If (fAbrRst(rs, Sql$)) Then
   mcianom = Trim(rs!razsoc)
   mciaregpat = rs!reg_pat
   mciaruc = rs!RUC
   mciadir = Trim(rs!direcc) & " " & rs!NRO & " " & rs!DIST
   mciadist = rs!DIST
End If

If rs.State = 1 Then rs.Close
mano = Val(Mid(Cmbfecha.Value, 7, 4))
mmes = Val(Mid(Cmbfecha.Value, 4, 2))
mesnombre = Name_Month(Format(mmes, "00"))
mlinea = 1
RUTA$ = App.Path & "\REPORTS\Boletas.txt"
Open RUTA$ For Output As #1
Barra.Max = Adocabeza.Recordset.RecordCount
If Opctotal.Value = True Then Adocabeza.Recordset.MoveFirst: Barra.Value = 0
If Opcindividual.Value <> True Then
   Panelprogress.Visible = True
   Panelprogress.ZOrder 0
   Me.Refresh
   If Opcrango.Value = True Then Barra.Value = Barra.Value = Adocabeza.Recordset.AbsolutePosition
End If

Do While Not Adocabeza.Recordset.EOF
   Barra.Value = Adocabeza.Recordset.AbsolutePosition
   If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
   Do While Not rsboleta.EOF
      rsboleta.Delete
      rsboleta.MoveNext
   Loop
   Print #1, Chr(15)
   mcad = Trim(mcianom)
   mcad = lentexto(65, Left(mcad, 65))
   Print #1, mcad & Space(4) & mcad
   'mCad = "REG. PAT. " & mciaregpat
   'mCad = lentexto(65, Left(mCad, 65))
   mcad = ""
   Print #1, mcad & Space(4) & mcad
   mciadir = lentexto(49, Left(mciadir, 49))
   mcad = mciadir & " RUC " & mciaruc
   mcad = lentexto(65, Left(mcad, 65))
   Print #1, mcad & Space(4) & mcad
   If VTipoPago = "04" Then
      mcad = "               PLANILLA EMPLEADOS - SUELDOS"
   Else
      mcad = "               PLANILLA OBREROS - SALARIOS"
   End If
   mcad = lentexto(65, Left(mcad, 65))
   Print #1, mcad & Space(4) & mcad
   Sql$ = nombre()
   If VTipoPago = "04" Then
      Sql$ = Sql$ & "a.*,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni " _
           & "from plahistorico a,planillas b " _
           & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & _
           VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " and a.placod='" & Adocabeza.Recordset!PlaCod & "' " _
           & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
           mperiodo = Format(mmes, "00")
   Else
      Sql$ = Sql$ & "a.*,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni " _
           & "from plahistorico a,planillas b " _
           & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and " _
           & "b.tipotrabajador='" & VTipotrab & "' and a.semana='" & Txtsemana.Text _
           & "' and year(a.fechaproceso)=" & mano & " and a.placod='" _
           & Adocabeza.Recordset!PlaCod & "' " _
           & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
           mperiodo = Txtsemana.Text
   End If
   If (fAbrRst(rs, Sql$)) Then
      mbasico = 0
      rs.MoveFirst
      
      Call llena_horas(rs)
      
      
      CRREPORTE.Sections(1).ReportObjects.Item("txtfec").SetText (rs("fechaproceso"))
      CRREPORTE.Sections(1).ReportObjects.Item("txtfec2").SetText (rs("fechaproceso"))
      CRREPORTE.Sections(1).ReportObjects.Item("TXTCOD").SetText (Trim(Adocabeza.Recordset!PlaCod))
      CRREPORTE.Sections(1).ReportObjects.Item("TXTCOD2").SetText (Trim(Adocabeza.Recordset!PlaCod))
      
      mnombre = lentexto(40, Left(rs!nombre, 40))
      CRREPORTE.Sections(1).ReportObjects.Item("TXTAPE").SetText (Trim(Left(mnombre, 30)))
      CRREPORTE.Sections(1).ReportObjects.Item("TXTAPE2").SetText (Trim(Left(mnombre, 28)))
      
      CRREPORTE.Sections(1).ReportObjects.Item("txttrem").SetText (rs("totaling"))
      CRREPORTE.Sections(1).ReportObjects.Item("txttrem2").SetText (rs("totaling"))
      
      tempo = rs("h01") + rs("h02") + rs("h03") + rs("h11") + rs("h10") + rs("h17")
      CRREPORTE.Sections(1).ReportObjects.Item("txttpag").SetText (tempo)
      CRREPORTE.Sections(1).ReportObjects.Item("txttpag2").SetText (tempo)
      
      
      Sql$ = "select importe,FACTOR_HORAS from plaremunbase where cia='" & wcia & "' and concepto='01' and placod='" & rs!PlaCod & "' and status<>'*'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then
         rs2.MoveFirst
         mbasico = rs2!importe: FACTOR_HORAS = rs2("FACTOR_HORAS")
      Else
         mbasico = 0
      End If
      CRREPORTE.Sections(1).ReportObjects.Item("TXTSUEL").SetText (mbasico)
      CRREPORTE.Sections(1).ReportObjects.Item("TXTSUEL2").SetText (mbasico)
      
      If rs2.State = 1 Then rs2.Close
      mcad = "NOMBRE    : " & mnombre & " DNI:" & Left(rs!dni, 8) & ")"
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(4) & mcad
      wciamae = Determina_Maestro_2("01055")
      Sql$ = "select cod_maestro3,descrip from maestros_3 where cod_maestro2='" & VTipotrab & "' and cod_maestro3='" & rs!cargo & "'"
      Sql$ = Sql$ & wciamae
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mcargo = rs2!DESCRIP Else mcargo = ""
      CRREPORTE.Sections(1).ReportObjects.Item("TXTCAT").SetText (mcargo)
      CRREPORTE.Sections(1).ReportObjects.Item("TXTCAT2").SetText (mcargo)
      
      If rs2.State = 1 Then rs2.Close
      mcargo = lentexto(31, Left(mcargo, 31))
      If Not IsNull(rs!fechacese) Then mcad = Format(rs!fechacese, "dd/mm/yyyy") Else mcad = Space(10)
      CRREPORTE.Sections(1).ReportObjects.Item("TXTFECCES").SetText (mcad)
      CRREPORTE.Sections(1).ReportObjects.Item("TXTFECCES2").SetText (mcad)
      
      mcad = "OCUPACION : " & mcargo & "BASICO : " & fCadNum(mbasico, "##,###,##0.00")
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(4) & mcad
      'If Not IsNull(RS!fechavacai) Then mcad = mcad & Format(RS!fechavacai, "dd/mm/yyyy") Else mcad = mcad & Space(10) & " "
      'If Not IsNull(RS!fechavacaf) Then mcad = mcad & Format(RS!fechavacaf, "dd/mm/yyyy") Else mcad = mcad & Space(10) & " "
      'If Not IsNull(RS!fechaproceso) Then mcad = mcad & Format(RS!fechaproceso, "dd/mm/yyyy") Else mcad = mcad & Space(10) & "   "
      mcad = "FECHA ING : " & Format(rs!fIngreso, "dd/mm/yyyy") & "                CARNET IPSS " & rs!ipss
      CRREPORTE.Sections(1).ReportObjects.Item("TXTFECIN").SetText (Format(rs!fIngreso, "dd/mm/yyyy"))
      CRREPORTE.Sections(1).ReportObjects.Item("TXTFECIN2").SetText (Format(rs!fIngreso, "dd/mm/yyyy"))
      
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(4) & mcad
      wciamae = Determina_Maestro("01069")
      Sql$ = "select * from maestros_2 where cod_maestro2='" & rs!CodAfp & "' and status<>'*'"
      Sql$ = Sql$ & wciamae
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mafp = rs2!DESCRIP Else mafp = ""
      If rs2.State = 1 Then rs2.Close
      mafp = lentexto(26, Left(mafp, 26))
      mcad = "AFP       : " & mafp & "CARNET SPP  " & rs!NUMAFP
      CRREPORTE.Sections(1).ReportObjects.Item("TXTAFP").SetText (Trim(mafp))
      CRREPORTE.Sections(1).ReportObjects.Item("TXTAFPP2").SetText (Trim(mafp))
      CRREPORTE.Sections(1).ReportObjects.Item("txtcus").SetText (rs!NUMAFP)
      CRREPORTE.Sections(1).ReportObjects.Item("txtcus2").SetText (rs!NUMAFP)
      
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(4) & mcad
      If VTipoPago = "04" Then
         mcad = "Mes       : " & mesnombre & "   " & Format(Right(Cmbfecha.Value, 4), "0000") & Space(5) & "(" & Trim(rs!PlaCod) & ")"
         CRREPORTE.Sections(1).ReportObjects.Item("TXTSEM").SetText (mesnombre)
         CRREPORTE.Sections(1).ReportObjects.Item("TXTSEM2").SetText (mesnombre)
      Else
         mcad = "SEMANA No : " & Txtsemana.Text & "  DEL " & Cmbdel.Value & " AL " & Cmbal.Value & Space(5) & "(" & Trim(rs!PlaCod) & ")"
         CRREPORTE.Sections(1).ReportObjects.Item("TXTSEM").SetText (Txtsemana.Text)
         CRREPORTE.Sections(1).ReportObjects.Item("TXTSEM2").SetText (Txtsemana.Text)
      End If
      
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(4) & mcad
      Print #1,
      mcad = "REMUNERACIONES:                DESCUENTOS:   Empleador Trabajador"
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(4) & mcad
      Print #1,
      
      'INGRESOS
      mCadSeteo = ""
      Sql$ = "select a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
           & "where a.cia='" & wcia & "' and a.status<>'*' and a.tipo='I' " _
           & "and a.cia=b.cia and b.status<>'*' and b.tipomovimiento='02' and a.tipo_trab='" & VTipotrab & "' and a.codigo=b.codinterno " _
           & "order by a.codigo"
      'Remunerativos
      If (fAbrRst(rs2, Sql$)) Then
         rs2.MoveFirst
         totremu = 0
         Do While Not rs2.EOF
            mbasico = 0
            Sql$ = "select * from plaafectos where cia='" & wcia & "' and status<>'*' and cod_remu='" & rs2!codigo & "'"
            If (fAbrRst(rsremu, Sql$)) Then
               mnumh = Remun_Horas(rs2!codigo)
               If mnumh <> 0 Then mbasico = rs(14 + mnumh)
               If mbasico <> 0 Then
                  mcad = lentexto(10, Left(rs2!Descripcion, 10)) & fCadNum(mbasico, "##0.00")
               Else
                  mcad = lentexto(16, Left(rs2!Descripcion, 16))
               End If
               mbasico = rs(44 + Val(rs2!codigo))
                              
               If rs2!codigo = "01" Then 'NORMAL
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTREMNOR").SetText (IIf(mbasico = 0, " ", mbasico))
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTREM2").SetText (IIf(mbasico = 0, " ", mbasico))
               End If
               If rs2!codigo = "02" Then 'ASIGNACION FAM.
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTREASF").SetText (IIf(mbasico = 0, " ", mbasico))
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTREASF2").SetText (IIf(mbasico = 0, " ", mbasico))
               End If
               If rs2!codigo = "03" Then 'ASIGN. MOVIL.
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTMOV").SetText (IIf(mbasico = 0, " ", mbasico))
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTMOV2").SetText (IIf(mbasico = 0, " ", mbasico))
               End If
               If rs2!codigo = "04" Then 'BON. T. SERVIC.
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTREMBOF").SetText (IIf(mbasico = 0, " ", mbasico))
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTrembof2").SetText (IIf(mbasico = 0, " ", mbasico))
               End If
               If rs2!codigo = "05" Then 'afp 10
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTafp10").SetText (IIf(mbasico = 0, " ", mbasico))
                  CRREPORTE.Sections(1).ReportObjects.Item("TXTafp102").SetText (IIf(mbasico = 0, " ", mbasico))
               End If
               
               Select Case rs2!codigo
               
               Case "06": 'afp 3
                    CRREPORTE.Sections(1).ReportObjects.Item("TXTafp3").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("TXTafp32").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "07":
               
               Case "08": 'sobretasa
                    CRREPORTE.Sections(1).ReportObjects.Item("txtrestasa").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtresta2").SetText (IIf(mbasico = 0, " ", mbasico))
                                   
               Case "09": 'dominical
                    CRREPORTE.Sections(1).ReportObjects.Item("txtredom").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtredom2").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "10": 'extras l-s
                    CRREPORTE.Sections(1).ReportObjects.Item("txtremls").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtremls2").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "11": 'extras d-f
                    CRREPORTE.Sections(1).ReportObjects.Item("txtedf").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtedf2").SetText (IIf(mbasico = 0, " ", mbasico))
               
               Case "12": 'feriados
                    CRREPORTE.Sections(1).ReportObjects.Item("txtremfe").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtremfe2").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "13": 'reintegros
                    CRREPORTE.Sections(1).ReportObjects.Item("txtrei").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtrei2").SetText (IIf(mbasico = 0, " ", mbasico))
               
               Case "14": 'vacaciones
                    CRREPORTE.Sections(1).ReportObjects.Item("txtrev").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtrev2").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "15": 'gratificacion
                    CRREPORTE.Sections(1).ReportObjects.Item("txtgrati").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtgrati2").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "16": 'otros pagos
                    CRREPORTE.Sections(1).ReportObjects.Item("txtot").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtot2").SetText (IIf(mbasico = 0, " ", mbasico))
                    
               Case "17": 'asignacion escolar
                    CRREPORTE.Sections(1).ReportObjects.Item("txtae").SetText (IIf(mbasico = 0, " ", mbasico))
                    CRREPORTE.Sections(1).ReportObjects.Item("txtae2").SetText (IIf(mbasico = 0, " ", mbasico))
               
               Case "21": '3ra ex ls
                      
                    tempo = Val(CRREPORTE.Sections(1).ReportObjects.Item("txtremls").Text)
                    CRREPORTE.Sections(1).ReportObjects.Item("txtremls").SetText (IIf(mbasico = 0, tempo, tempo + mbasico))
                    'tempo = Val(CRREPORTE.Sections(1).ReportObjects.Item("txtremls2").Text)
                    CRREPORTE.Sections(1).ReportObjects.Item("txtremls2").SetText (IIf(mbasico = 0, tempo, mbasico + tempo))
                                  
               End Select
               
               If mbasico <> 0 Then
                  mcad = mcad & Space(1) & fCadNum(mbasico, "###,##0.00")
               Else
                  mcad = mcad & Space(1) & Space(10)
               End If
               rsboleta.AddNew
               rsboleta!texto = mcad
               totremu = totremu + mbasico
            Else
'               Debug.Print "ME VOY"
            End If
            If rsremu.State = 1 Then rsremu.Close
            mCadSeteo = mCadSeteo & rs2!codigo
            
            If rs2("CODIGO") = "27" Then
               'CALCULA PROM/DIARIO
               con = "SELECT CAST( (" & totremu & "/(h01+h02+h03+h04+h05+h12)) AS DECIMAL(5,2))*" & FACTOR_HORAS & " AS PROM" & " FROM " & _
               "PLAHISTORICO WHERE PLACOD='" & Trim(Adocabeza.Recordset!PlaCod) & "' AND " & _
               "YEAR(FECHAPROCESO)=" & mano & " AND month(fechaproceso)=" & mmes & " AND " & _
               "SEMANA='" & Trim(Txtsemana) & "'"
               RX.Open con, cn, adOpenStatic, adLockReadOnly
                  rsboleta.AddNew
                  rsboleta("texto") = "PROM/DIARIO" & Space(1) & Space(10) & RX(0)
                  
               CRREPORTE.Sections(1).ReportObjects.Item("txtprdi").SetText (RX(0))
               CRREPORTE.Sections(1).ReportObjects.Item("txtprdi2").SetText (RX(0))
                                 
                   'totremu = totremu + RX(0)
               RX.Close
            End If
            
            rs2.MoveNext
         Loop
             
      End If
      
      totremu = totremu + No_Seteados("IN", mCadSeteo, "R")
      rsboleta.AddNew
      rsboleta!texto = Space(17) & "----------"
      rsboleta.AddNew
      rsboleta!texto = "Sub Total I        " & Format(totremu, "###,###,##0.00")
      rsboleta.AddNew
      rsboleta!texto = ""
      rsboleta.AddNew
      rsboleta!texto = "NO REMUNERATIVOS :"
      rsboleta.AddNew
      rsboleta!texto = ""
      
      'No Remunerativos
      mCadSeteo = ""
      If rs2.RecordCount > 0 Then rs2.MoveFirst
      totremu = 0
      Do While Not rs2.EOF
         mbasico = 0
         Sql$ = "select * from plaafectos where cia='" & wcia & "' and status<>'*' and cod_remu='" & rs2!codigo & "'"
         If Not (fAbrRst(rsremu, Sql$)) Then
            mnumh = Remun_Horas(rs2!codigo)
            If mnumh <> 0 Then mbasico = rs(14 + mnumh)
            If mbasico <> 0 Then
               mcad = lentexto(11, Left(rs2!Descripcion, 11)) & fCadNum(mbasico, "#0.00")
            Else
               mcad = lentexto(16, Left(rs2!Descripcion, 16))
            End If
            mbasico = rs(44 + Val(rs2!codigo))
            If mbasico <> 0 Then
               mcad = mcad & Space(1) & fCadNum(mbasico, "###,##0.00")
            Else
               mcad = mcad & Space(1) & Space(10)
            End If
            rsboleta.AddNew
            rsboleta!texto = mcad
            totremu = totremu + mbasico
         End If
         If rsremu.State = 1 Then rsremu.Close
         mCadSeteo = mCadSeteo & rs2!codigo
         rs2.MoveNext
      Loop
      
      totremu = totremu + No_Seteados("IN", mCadSeteo, "N")
      rsboleta.AddNew
      rsboleta!texto = Space(17) & "----------"
      rsboleta.AddNew
      If totremu = 0 Then
         rsboleta!texto = "Sub Total II           " & Format(totremu, "###,###,##0.00")
      Else
         rsboleta!texto = "Sub Total II         " & Format(totremu, "###,###,##0.00")
      End If
      
      'APORTACIONES
      mCadSeteo = ""
      Sql$ = "select a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
           & "where a.cia='" & wcia & "' and a.status<>'*' and a.tipo='A' " _
           & "and a.cia=b.cia and b.status<>'*' and a.tipo_trab='" & VTipotrab & "' and b.tipomovimiento='03' and a.codigo=b.codinterno " _
           & "order by a.codigo"
      If (fAbrRst(rs2, Sql$)) Then
         mlinea = 1
         Do While Not rs2.EOF
            mbasico = rs(44 + 50 + 20 + Val(rs2!codigo))
            
            If rs2!codigo = "01" Then 'esalud
               CRREPORTE.Sections(1).ReportObjects.Item("txtaess").SetText (IIf(mbasico = 0, " ", mbasico))
               CRREPORTE.Sections(1).ReportObjects.Item("txtaess2").SetText (IIf(mbasico = 0, " ", mbasico))
            End If
            
            If rs2!codigo = "03" Then 'senati
               CRREPORTE.Sections(1).ReportObjects.Item("txtsenati").SetText (IIf(mbasico = 0, " ", mbasico))
               CRREPORTE.Sections(1).ReportObjects.Item("txtsenati2").SetText (IIf(mbasico = 0, " ", mbasico))
            End If
            
            If mbasico <> 0 Then
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(1) & fCadNum(mbasico, "##,##0.00")
            Else
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(1) & Space(9)
            End If
            rsboleta.AbsolutePosition = mlinea
            rsboleta!texto = Left(rsboleta!texto, 27) & Space(4) & mcad
            mlinea = mlinea + 1
            mCadSeteo = mCadSeteo & rs2!codigo
            rs2.MoveNext
         Loop
         
         'descuentos
         CRREPORTE.Sections(1).ReportObjects.Item("txtapat").SetText (IIf(rs("totalapo") = 0, " ", rs("totalapo")))
         CRREPORTE.Sections(1).ReportObjects.Item("txtappat2").SetText (IIf(rs("totalapo") = 0, " ", rs("totalapo")))
      End If
      If rs2.State = 1 Then rs2.Close
      Call No_Seteados("AP", mCadSeteo, "")
      'DESCUENTOS"
      mCadSeteo = ""
      Sql$ = "select a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
           & "where a.cia='" & wcia & "' and a.status<>'*' and a.tipo='D' " _
           & "and a.cia=b.cia and b.status<>'*' and a.tipo_trab='" & VTipotrab & "' and b.tipomovimiento='03' and a.codigo=b.codinterno " _
           & "order by a.codigo"
      If (fAbrRst(rs2, Sql$)) Then
         Do While Not rs2.EOF
            mbasico = rs(44 + 50 + Val(rs2!codigo))
            
            Select Case rs2!codigo
            
            Case "04": 'snp
                 CRREPORTE.Sections(1).ReportObjects.Item("txtasnp").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtasnp2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "05": 'judicial
                 CRREPORTE.Sections(1).ReportObjects.Item("txtju").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtadju2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "06": 'essalud vida
                 CRREPORTE.Sections(1).ReportObjects.Item("txtes").SetText ((IIf(mbasico = 0, " ", mbasico)))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtes2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "07" 'cta. cte
                 
                 con = "select numinterno from plabolcte where cia='" & wcia & "' and " & _
                 "placod='" & Trim(Adocabeza.Recordset!PlaCod) & "' and year(fechaproceso)='" & _
                  mano & "' and month(fechaproceso)='" & mmes & "' and " & _
                  "day(fechaproceso)='" & Format(rs(FechaProceso), "dd") & "' and status<>'*'"
                 RX.Open con, cn, adOpenStatic, adLockReadOnly
                   If RX.RecordCount > 0 Then cod = RX("numinterno")
                 RX.Close
                 
                 If Trim(cod) <> "" Then
                    fecha = Trim(Txtsemana) & "/" & mmes & "/" & mano
                    con = "set dateformat dmy "
                    con = con & "select sum(importe) as tot from plabolcte where cia='" & wcia & "'" & _
                    " and placod='" & Trim(Adocabeza.Recordset!PlaCod) & "' and " & _
                    " numinterno='" & cod & "' and fechaproceso<'" & fecha & "' and " & _
                    "status<>'*' "
                 
                    RX.Open con, cn, adOpenStatic, adLockReadOnly
                       If RX.RecordCount > 0 Then
                          CRREPORTE.Sections(1).ReportObjects.Item("txtsant").SetText (RX("tot"))
                          CRREPORTE.Sections(1).ReportObjects.Item("txtsant2").SetText (RX("tot"))
                       End If
                    RX.Close
                 End If
            
                 CRREPORTE.Sections(1).ReportObjects.Item("txtcta").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtadcta2").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtabon").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtabon2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "08": 'sindicato
                 CRREPORTE.Sections(1).ReportObjects.Item("txtsin").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtsin2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "09": 'adelanto de quincena
                 CRREPORTE.Sections(1).ReportObjects.Item("txtad").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtad2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "10": 'tardanzas
            
            Case "11": 'afp
                 CRREPORTE.Sections(1).ReportObjects.Item("txttafp").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txtafp2").SetText (IIf(mbasico = 0, " ", mbasico))
            Case "12": 'otros descuentos
                 CRREPORTE.Sections(1).ReportObjects.Item("txtotdes").SetText (IIf(mbasico = 0, " ", mbasico))
                 CRREPORTE.Sections(1).ReportObjects.Item("txttotdes2").SetText (IIf(mbasico = 0, " ", mbasico))
                 
            End Select
                   
            If mbasico <> 0 Then
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(13) & fCadNum(mbasico, "##,##0.00")
            Else
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(13) & Space(9)
            End If
            rsboleta.AbsolutePosition = mlinea
            mCadBlanc = ""
            If Len(RTrim(rsboleta!texto)) < 27 Then
               For I = Len(RTrim((rsboleta!texto))) To 26
                   mCadBlanc = mCadBlanc & Space(1)
               Next
               rsboleta!texto = RTrim(rsboleta!texto) & mCadBlanc & Space(4) & mcad
            Else
               rsboleta!texto = Left(rsboleta!texto, 27) & Space(4) & mcad
            End If
            mlinea = mlinea + 1
            mCadSeteo = mCadSeteo & rs2!codigo
            rs2.MoveNext
         Loop
      End If
      Call No_Seteados("DE", mCadSeteo, "")
      If rs2.State = 1 Then rs2.Close
   End If
   If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
   Do While Not rsboleta.EOF
      mtexto = lentextosp(65, Left(rsboleta!texto, 65))
      Print #1, mtexto & Space(4) & mtexto
      rsboleta.MoveNext
   Loop
   Print #1, Space(18) & String(47, "-") & Space(4) & Space(18) & String(47, "-")
   mtexto = "*** TOTAL *** " & fCadNum(rs!totaling, "##,###,##0.00") & "   * TOTAL *    " & fCadNum(rs!totalapo, "###,##0.00") & "  " & fCadNum(rs!totalded, "###,##0.00")
   With CRREPORTE.Sections(1).ReportObjects
   
         .Item("txttrem").SetText (Trim(fCadNum(rs!totaling, "##,###,##0.00")))
         .Item("txttrem2").SetText (.Item("txttrem").Text)
         .Item("txtapat").SetText (Trim(fCadNum(rs!totalapo, "###,##0.00")))
         .Item("txtappat2").SetText (.Item("txtapat").Text)
         .Item("txtdes").SetText (Trim(fCadNum(rs!totalded, "###,##0.00")))
         .Item("txtdes2").SetText (.Item("txtdes").Text)
        
   End With
   
   mtexto = lentexto(65, Left(mtexto, 65))
   Print #1, mtexto & Space(4) & mtexto
   Print #1, String(65, "-") & Space(4) & String(65, "-")
   Print #1, Space(34) & "NETO A PAGAR      " & fCadNum(rs!totaling - rs!totalded, "##,###,##0.00") & Space(4) & Space(34) & "NETO A PAGAR      " & fCadNum(rs!totaling - rs!totalded, "##,###,##0.00")
   CRREPORTE.Sections(1).ReportObjects.Item("txtneto").SetText (Trim(fCadNum(rs!totaling - rs!totalded, "##,###,##0.00")))
   CRREPORTE.Sections(1).ReportObjects.Item("txtneto2").SetText (Trim(fCadNum(rs!totaling - rs!totalded, "##,###,##0.00")))
   Print #1,
   Print #1,
   Print #1,
   mtexto = "FECHA DE PAGO : " & Format(Left(Cmbfecha.Value, 2), "00") & " DE " & mesnombre & " DE " & Format(mano, "0000")
   mtexto = lentexto(65, Left(mtexto, 65))
   Print #1, mtexto & Space(4) & mtexto
   mtexto = Left(mcianom, 40) & "             ---------------------"
   mtexto = lentexto(65, Left(mtexto, 65))
   Print #1, mtexto & Space(4) & mtexto
   Print #1, Space(47) & "Recibi Conforme" & Space(4) & Space(50) & "Recibi Conforme"
   
   If rs.State = 1 Then rs.Close
   If Opcindividual.Value = True Then Exit Do
   Adocabeza.Recordset.MoveNext
   Print #1, SaltaPag
Loop
Close #1
Panelprogress.Visible = False
FramePrint.Visible = False

CRV.ReportSource = CRREPORTE
CRV.ViewReport
'Call Imprime_Txt("Boletas.txt", ruta$)
Exit Sub

FUNKA:
Close #1
MsgBox "Error : " & ERR.Description, vbCritical, "Planillas"
End Sub

Private Sub Print_Quin()
Dim mano As Integer
Dim mmes As Integer
Dim mnombre As String
Dim wciamae As String
Dim mesnombre As String
Dim mcianom As String
Dim mciadep As String
Dim montolet As String


'Sql$ = "select a.*,dist,b.dpto as dptou from cia a,ubigeos b where cod_cia='" & wcia & "' and a.cod_ubi=b.cod_ubi"
Sql$ = " select c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dptou, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais from cia c, sunat_ubigeo u " & _
      "  WHERE cod_cia='" & wcia & "' AND u.id_ubigeo=*c.cod_ubi"

If (fAbrRst(rs, Sql$)) Then
   mcianom = Trim(rs!razsoc)
   mciadep = Trim(rs!dptou)
End If
rs.Close
mesnombre = Name_Month(Format(Cmbfecha.Month, "00"))
mano = Val(Mid(Cmbfecha.Value, 7, 4))

RUTA$ = App.Path & "\REPORTS\Quincena.txt"
Open RUTA$ For Output As #1
If Opctotal.Value = True Then Adocabeza.Recordset.MoveFirst
Do While Not Adocabeza.Recordset.EOF
   If Adocabeza.Recordset!totneto > 0 Then montolet = monto_palabras(Adocabeza.Recordset!totneto) Else montolet = ""
   Sql$ = nombre()
   Sql$ = Sql$ & "cargo " _
           & "from planillas " _
           & "where cia='" & wcia & "' and placod='" & Adocabeza.Recordset!PlaCod & "' " _
           & "and status<>'*' "
           mperiodo = Txtsemana.Text
   If (fAbrRst(rs, Sql$)) Then mnombre = lentexto(65, Left(rs!nombre, 65))
   rs.Close

   Print #1,
   Print #1,
   Print #1, Chr(18) & Space(40) & "NETO A PAGAR S/.       " & fCadNum(Adocabeza.Recordset!totneto, "#,###,###.00")
   Print #1,
   Print #1,
   Print #1, Chr(15) & Space(5) & "RECIBI DE LA CIA.  " & Chr(18) & mcianom & Chr(15)
   Print #1,
   Print #1, Space(5) & Chr(18) & "LA CANTIDAD DE           " & Chr(15) & AsteriscoR(80, montolet)
   Print #1,
   Print #1,
   Print #1, Space(5) & Chr(18) & "POR CONCEPTO DE           " & Chr(15) & "Pago de la 1ra Quincena del mes de " & mesnombre & " de " & Format(mano, "0000") & Chr(18)
   Print #1,
   Print #1,
   Print #1,
   Print #1,
   Print #1, Space(50) & mciadep & " " & Format(Cmbfecha.Day, "00") & " DE " & mesnombre & " DE " & Format(mano, "0000")
   Print #1,
   Print #1,
   Print #1,
   Print #1,
   Print #1, Chr(15) & Space(87) & "-------------------------------------"
   Print #1, Space(87) & "                FIRMA"
   Print #1,
   Print #1, Space(5) & Chr(18) & "NOMBRE        " & Chr(15) & mnombre
   Print #1,
   Print #1,
   Print #1,
   Print #1,
   Print #1, Space(50) & "-------------------------------------"
   Print #1, Space(50) & mcianom
   If Opcindividual.Value = True Then Exit Do
   Adocabeza.Recordset.MoveNext
Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("Quincena.txt", RUTA$)
End Sub

Sub llena_horas(rs As ADODB.Recordset)
    Dim I As Integer

    With CRREPORTE.Sections(1).ReportObjects
    
         'For i = 1 To 10
             .Item("txtnom").SetText (IIf(rs("h01") = 0, " ", rs("H01")))
             .Item("txtnom2").SetText (IIf(rs("h01") = 0, " ", rs("H01")))
             
             .Item("txtdom").SetText (IIf(rs("h02") = 0, " ", rs("H02")))
             .Item("txtdom2").SetText (IIf(rs("h02") = 0, " ", rs("H02")))
             
             .Item("txtfe").SetText (IIf(rs("h03") = 0, " ", rs("H03")))
             .Item("txtfer2").SetText (IIf(rs("h03") = 0, " ", rs("H03")))
             
             .Item("txtppag").SetText (IIf(rs("h04") = 0, " ", rs("H04")))
             .Item("txtppag2").SetText (IIf(rs("h04") = 0, " ", rs("H04")))
             
             .Item("txtenfpag").SetText (IIf(rs("h05") = 0, " ", rs("H05")))
             .Item("txtenfpag2").SetText (IIf(rs("h05") = 0, " ", rs("H05")))
             
             .Item("txtenfnopag").SetText (IIf(rs("h06") = 0, " ", rs("H06")))
             .Item("txtenfnopag2").SetText (IIf(rs("h06") = 0, " ", rs("H06")))
             
             .Item("txtacct").SetText (IIf(rs("h07") = 0, " ", rs("H07")))
             .Item("txtacct2").SetText (IIf(rs("h07") = 0, " ", rs("H07")))
             
             .Item("txtfainj").SetText (IIf(rs("h08") = 0, " ", rs("H08")))
             .Item("txtfainj2").SetText (IIf(rs("h08") = 0, " ", rs("H08")))
             
             .Item("txtsus").SetText (IIf(rs("h09") = 0, " ", rs("H09")))
             .Item("txtsus2").SetText (IIf(rs("h09") = 0, " ", rs("H09")))
             
             '====================================
             'horas extras ls
             .Item("txtls").SetText (IIf(rs("h10") + rs("h17") = 0, " ", rs("h10") + rs("h17")))
             .Item("txtls2").SetText (IIf(rs("h10") + rs("h17") = 0, " ", rs("h10") + rs("h17")))
             
             'horas extras df
             .Item("txtdf").SetText (IIf(rs("h11") = 0, " ", rs("H11")))
             .Item("txtdf2").SetText (IIf(rs("h11") = 0, " ", rs("H11")))
             
             'vacaciones
             .Item("txtvaca").SetText (IIf(rs("h12") = 0, " ", rs("H12")))
             .Item("txtvaca2").SetText (IIf(rs("h12") = 0, " ", rs("H12")))
             
             'sobretasa
             .Item("txtst").SetText (IIf(rs("h13") = 0, " ", rs("H13")))
             .Item("txtst2").SetText (IIf(rs("h13") = 0, " ", rs("H13")))
             
             'dias trabajados
             .Item("txtdt").SetText (IIf(rs("h14") = 0, " ", rs("H14")))
             .Item("txtdt2").SetText (IIf(rs("h14") = 0, " ", rs("H14")))
             
             'otros
             .Item("txthot").SetText (IIf(rs("h15") = 0, " ", rs("H15")))
             .Item("txthot2").SetText (IIf(rs("h15") = 0, " ", rs("H15")))
             
              temp = rs("h01") + rs("h02") + rs("h03") + rs("h04") + rs("h05") + rs("h06")
              temp = temp + rs("h07") + rs("h08") + rs("h09") + rs("h10") + rs("h17")
              temp = temp + rs("h11") + rs("h12") + rs("h13") + rs("h14") + rs("h15")
              
             'total pagadas
             .Item("txttpag").SetText (IIf(temp = 0, " ", temp))
             .Item("txttpag2").SetText (IIf(temp = 0, " ", temp))
             
         'Next
    
    End With
End Sub

Sub Limpiar()
    With CRREPORTE.Sections(1).ReportObjects
          .Item("txtpnp").SetText ("")
          .Item("txtpnp2").SetText ("")
          .Item("txtess").SetText ("")
          .Item("txtvac").SetText ("")
          .Item("txthas").SetText ("")
          .Item("txtfec").SetText ("")
          .Item("txtper").SetText ("")
          .Item("TXTPER2").SetText ("")
          .Item("txtutil").SetText ("")
          .Item("txtutil2").SetText ("")
          .Item("txtaes").SetText ("")
          .Item("txtaes2").SetText ("")
          .Item("txt5cat").SetText ("")
          .Item("txt5cat2").SetText ("")
          .Item("txtsolsind").SetText ("")
          .Item("txtsolsi2").SetText ("")
          .Item("txttafp").SetText ("")
          .Item("txtafp2").SetText ("")
          .Item("TXTAESS").SetText ("")
          .Item("txtaess2").SetText ("")
          .Item("txtaesnp").SetText ("")
          .Item("txtaesss2").SetText ("")
          .Item("txtaeies").SetText ("")
          .Item("txtaeies2").SetText ("")
          .Item("txtaesc").SetText ("")
          .Item("txtades2").SetText ("")
          .Item("txtdes").SetText ("")
          .Item("txtdes2").SetText ("")
          .Item("txtsenati").SetText ("")
          .Item("txtsenati2").SetText ("")
          .Item("txtsant").SetText ("")
          .Item("txtsant2").SetText ("")
          .Item("txtcar").SetText ("")
          .Item("txtcar2").SetText ("")
          .Item("txtabon").SetText ("")
          .Item("txtabon2").SetText ("")
          .Item("txtsact").SetText ("")
          .Item("txtsact2").SetText ("")
          .Item("txtneto").SetText ("")
          .Item("txtneto2").SetText ("")
          .Item("txtcus").SetText ("")
          .Item("txtcus2").SetText ("")
          .Item("txtes").SetText ("")
          .Item("txtess2").SetText ("")
          .Item("txtes2").SetText ("")
    End With

End Sub
