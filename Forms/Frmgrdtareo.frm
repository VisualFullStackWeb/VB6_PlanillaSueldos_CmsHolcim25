VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmgrdtareo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Tareo"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "Frmgrdtareo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbccosto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1500
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Txtfechaa 
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Txtfechade 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Txtcodobra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Txtcodtrab 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5175
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   9135
      Begin MSDataGridLib.DataGrid GrdTareo 
         Bindings        =   "Frmgrdtareo.frx":030A
         Height          =   4935
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8705
         _Version        =   393216
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
            DataField       =   "fecha"
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
         BeginProperty Column01 
            DataField       =   "nombre"
            Caption         =   "Trabajador"
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
            DataField       =   "codigotrab"
            Caption         =   "codigotrab"
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
            DataField       =   "obra"
            Caption         =   "obra"
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
            DataField       =   "ccosto"
            Caption         =   "ccosto"
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
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   6974.93
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoTareo 
         Height          =   375
         Left            =   720
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   7560
         TabIndex        =   3
         Top             =   120
         Width           =   1455
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
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Lbltipo 
      Height          =   135
      Left            =   0
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "C. de Costo"
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "A"
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "De"
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label Lblobra 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Lblnombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Obra"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "Frmgrdtareo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VArea As String
Private Sub Cmbccosto_Click()

If Trim(TxtCodTrab.Text) = "" Then
   MsgBox "Ingrese Codigo del Trabajador", vbInformation, Me.Caption
   Exit Sub
End If

If CmbCcosto.Text = "TODOS" Then
   VArea = ""
Else
   VArea = fc_CodigoComboBox(CmbCcosto, 3)
End If
Call Procesa_Tareo
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01044", "", CmbCcosto)
CmbCcosto.AddItem ("TODOS")
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 9225
Me.Height = 7470
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
LblFecha.Caption = Date
Txtfechade.Text = Format(Date, "dd/mm/yyyy")
Txtfechaa.Text = Format(Date, "dd/mm/yyyy")
Procesa_Tareo
End Sub

Private Sub GrdTareo_DblClick()
If AdoTareo.Recordset.RecordCount <= 0 Then Exit Sub
Call Frmtareo.Carga_Tareo(Trim(GrdTareo.Columns(2)), Trim(GrdTareo.Columns(3)), Trim(GrdTareo.Columns(0)), GrdTareo.Columns(4))
End Sub

Private Sub Txtcodobra_GotFocus()
wbus = "OB"
NameForm = "Frmgrdtareo"
End Sub

Private Sub Txtcodobra_KeyPress(KeyAscii As Integer)
Txtcodobra.Text = Txtcodobra.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Txtcodobra.Text = Format(Txtcodobra.Text, "00000000"): Txtfechade.SetFocus
End Sub

Private Sub Txtcodobra_LostFocus()
wbus = ""
If Txtcodobra.Text <> "" Then
   Sql$ = "select cod_obra,descrip,status from plaobras where cod_cia='" & wcia & "' and cod_obra='" & Txtcodobra.Text & "' order by status"
   If (fAbrRst(rs, Sql$)) Then
      If rs!status = "*" Then
         MsgBox "Obra Eliminada", vbInformation, "Registro de Personal"
         Lblobra.Caption = ""
         Txtcodobra.SetFocus
      Else
         Lblobra.Caption = Trim(rs!DESCRIP)
      End If
   Else
     MsgBox "Codigo de Obra no Registrada", vbInformation, "Registro de Personal"
     Lblobra.Caption = ""
     Txtcodobra.SetFocus
   End If
End If
Procesa_Tareo
End Sub

Private Sub TxtCodTrab_GotFocus()
wbus = "PL"
End Sub

Private Sub TxtCodTrab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Txtfechade.SetFocus
End If
End Sub
Private Sub TxtCodTrab_LostFocus()
Dim mcod As String
On Error GoTo FUNKA
If Trim(TxtCodTrab.Text) <> "" Then
    Sql$ = nombre()
    Sql$ = Sql$ + "tipotrabajador,status from planillas where cia='" & wcia & "' AND placod='" & TxtCodTrab.Text & "' order by status"
   
    cn.CursorLocation = adUseClient
    Set rs = New ADODB.Recordset
    Set rs = cn.Execute(Sql$)
    If rs.RecordCount > 0 Then
       If rs!status = "*" Then
          MsgBox "Trabajador Eliminado", vbExclamation, "Codigo N° => " & TxtCodTrab.Text
          LblNombre.Caption = ""
          Lbltipo.Caption = ""
          TxtCodTrab.SetFocus
       Else
          LblNombre.Caption = rs!nombre
          Lbltipo.Caption = rs!TipoTrabajador
       End If
    Else
       MsgBox "Codigo de Trabajador no Registrado", vbExclamation, "Codigo N° => " & TxtCodTrab.Text
       LblNombre.Caption = ""
       Lbltipo.Caption = ""
       TxtCodTrab.SetFocus
    End If
Else
   LblNombre.Caption = ""
   Lbltipo.Caption = ""
End If
wbus = ""
If Lbltipo.Caption = "05" Then
   Label5.Visible = True
   Txtcodobra.Visible = True
   Lblobra.Visible = True
Else
   Label5.Visible = False
   Txtcodobra.Visible = False
   Lblobra.Visible = False
End If
Procesa_Tareo
Exit Sub
FUNKA:
   MsgBox "Error: " & ERR.Description, vbCritical, "Planillas"
End Sub
Public Sub Procesa_Tareo()
Dim cad As String
Dim cadF As String
Dim cad2 As String
Dim cadF2 As String
Dim FEC As String
cad = ""
cadF = ""
cad2 = ""
cadF2 = ""
Select Case StrConv(gsAdminDB, 1)
       Case Is = "MYSQL"
            FEC = "concat(dayofmonth(fecha),'/', month(fecha),'/',year(fecha)) as fecha, " _
                & "CONCAT(rtrim(ap_pat), ' ',rtrim(ap_mat),' ',rtrim(nom_1),' ',rtrim(nom_2)) AS nombre "
       Case Is = "SQL SERVER"
'            FEC = "rtrim(str(day(Fecha)))+'/'+ltrim(str(month(Fecha)))+'/'+ltrim(str(year(Fecha))) as fecha, " _
'                & "rtrim(ap_pat)+' '+rtrim(ap_mat)+' '+rtrim(nom_1)+' '+rtrim(nom_2) as nombre "
            FEC = "convert(datetime,convert(varchar(10),fecha,103),103) as fecha, " _
                & "rtrim(ap_pat)+' '+rtrim(ap_mat)+' '+rtrim(nom_1)+' '+rtrim(nom_2) as nombre "
End Select

If IsDate(Txtfechade.Text) And Not IsDate(Txtfechaa.Text) Then cad = " and Fecha >= ": cadF = Format(Txtfechade.Text, FormatFecha) + FormatTimei
If Not IsDate(Txtfechade.Text) And IsDate(Txtfechaa.Text) Then cad = "and Fecha <= ": cadF = Format(Txtfechaa.Text, FormatFecha) + FormatTimef
If IsDate(Txtfechade.Text) And IsDate(Txtfechaa.Text) Then cad = "and fecha BETWEEN ": cadF = Format(Txtfechade.Text, FormatFecha) & FormatTimei: cad2 = "AND ": cadF2 = Format(Txtfechaa.Text, FormatFecha) & FormatTimef
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "select distinct(codigotrab),a.obra,a.ccosto,a.cia,"
Sql$ = Sql$ & FEC
Sql$ = Sql$ & "from platareo a,planillas b " _
     & "where a.cia='01' and a.cia=b.cia and a.status<>'*' and a.codigotrab=b.placod " _
     & "and codigotrab like '" & Trim(TxtCodTrab.Text) + "%" & "' and a.obra like '" & Trim(Txtcodobra.Text) + "%" & "' " _
     & "and ccosto like '" & Trim(VArea) + "%" & "' "
If cad <> "" Then Sql$ = Sql$ & cad
If cadF <> "" Then Sql$ = Sql$ & "'" & cadF & "'"
If cad2 <> "" Then Sql$ = Sql$ & cad2
If cadF2 <> "" Then Sql$ = Sql$ & "'" & cadF2 & "'"
Sql$ = Sql$ & " order by a.fecha"
cn.CursorLocation = adUseClient
Set AdoTareo.Recordset = cn.Execute(Sql$, 64)

GrdTareo.Refresh
End Sub

Private Sub Txtfechaa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtfechade.SetFocus
End Sub

Private Sub Txtfechaa_LostFocus()
If Txtfechaa.Text <> "__/__/____" And Not IsDate(Txtfechaa.Text) Then
   MsgBox "INGRESE FECHA CORRECTAMENTE", vbInformation, "Tareo"
   Txtfechaa.SetFocus
Else
   Procesa_Tareo
End If

End Sub

Private Sub Txtfechade_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtfechaa.SetFocus
End Sub

Private Sub Txtfechade_LostFocus()
If Trim(Txtfechade.Text) <> "__/__/____" And Not IsDate(Txtfechade.Text) Then
   MsgBox "INGRESE FECHA CORRECTAMENTE", vbInformation, "Tareo"
   Txtfechade.SetFocus
Else
   Procesa_Tareo
End If

End Sub
