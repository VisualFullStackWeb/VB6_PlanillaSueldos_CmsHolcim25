VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmprovision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provisiones"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "Frmprovision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   960
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1296
      _StockProps     =   15
      ForeColor       =   8388608
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6300
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   75
         Width           =   6135
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
         TabIndex        =   9
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5310
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   7455
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "Frmprovision.frx":030A
         Height          =   5160
         Left            =   75
         TabIndex        =   6
         Top             =   75
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   9102
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
         ColumnCount     =   3
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "provmes"
            Caption         =   "Provision"
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
         BeginProperty Column02 
            DataField       =   "placod"
            Caption         =   "codigo"
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
               ColumnWidth     =   5715.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoCabeza 
         Height          =   330
         Left            =   1200
         Top             =   3000
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   465
      Width           =   7455
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   75
         Width           =   2175
      End
      Begin VB.TextBox Txtano 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   75
         Width           =   495
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "Frmprovision.frx":0322
         Left            =   840
         List            =   "Frmprovision.frx":034A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   75
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   660
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   330
         Left            =   3480
         TabIndex        =   3
         Top             =   90
         Width           =   255
         Size            =   "450;582"
      End
   End
   Begin VB.Label Lbltipo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3240
      TabIndex        =   15
      Top             =   6300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   6300
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Provision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4560
      TabIndex        =   10
      Top             =   6300
      Width           =   1290
   End
End
Attribute VB_Name = "Frmprovision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nFil As Integer
Dim nCol As Integer
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim VTipo As String
Dim mHourMonth As Integer
Dim rs2 As ADODB.Recordset
Dim SQL As String

Dim ArrReporte() As Variant

Const COL_CODIGO = 0
Const COL_NOMBRE = 1
Const COL_RCDAÑOANTERIOR = 2
Const COL_VACACTOMADAS = 3
Const COL_RCDAÑOACTUAL = 4
Const COL_RCDACUMULADO = 5

Const HORAS_EMPLEADO = 240
Const HORAS_OBRERO = 48
Const hORAS_X_DIA = 8
Const DIAS_TRABAJO = 30


Private Sub Cmbcia_Click()
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Cmbtipo.AddItem "TOTAL"
Cmbtipo.ItemData(Cmbtipo.NewIndex) = "99"
End Sub

Private Sub Cmbmes_Click()
Procesa_Consultas
End Sub

Private Sub Cmbtipo_Click()
Dim wciamae As String
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
wciamae = Determina_Maestro("01076")

If VTipo = "01" Then
   SQL = "Select flag2 from maestros_2 where cod_maestro2='04' and status<>'*'"
Else
   SQL = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
End If
SQL = SQL$ & wciamae
mHourMonth = 0
If (fAbrRst(rs, SQL)) Then mHourMonth = Val(rs!flag2)
rs.Close
Procesa_Consultas
End Sub

Private Sub Command1_Click()
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If Lbltipo.Caption = "V" Then Reporte_Provision_Vaca
If Lbltipo.Caption = "G" Then Reporte_Provision_Grati
End Sub

Private Sub Command4_Click()
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If Lbltipo.Caption = "V" Then PROVISION_VACACIONES ' Calcula_Provision_Vaca
If Lbltipo.Caption = "G" Then PROVICIONES_GRATI 'Calcula_Provision_Grati
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7125
Me.Width = 7530
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Txtano.Text = Format(Year(Date), "0000")
Cmbmes.ListIndex = Month(Date) - 1
End Sub
Private Sub Calcula_Provision_Vaca()
Dim dia As Integer
Dim fecproc As String
Dim fing As String
Dim fpromed As String
Dim mfecanoant As String
Dim raAnt As String
Dim raAct As String
Dim rVaca As String
Dim rAcum As String
Dim VacaPag As Integer
Dim mFactProm As Integer
Dim cont1 As Integer
Dim mCadIng As String
Dim mBase As Currency
Dim mtoting As Currency
Dim mCadProm As String
Dim mfectope As String
Dim i As Integer
Dim j As Integer
Dim X As Integer
Dim nFields As Integer
Dim mProvAnt As Currency
Dim mVacPagada As Currency
Dim mVacPorPagar As Currency
Dim mProvMes As Currency
Dim PLAS(0 To 50) As Double

dia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
fecproc = Format(dia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Trim(Txtano.Text)
mfecanoant = "31/12/" & Format(Val(Txtano.Text) - 1, "0000")

SQL = "select * from planillas where cia='" & wcia & "' " & _
"and tipotrabajador='" & VTipo & "' and fcese is null and status<>'*' order by placod"

If (fAbrRst(rs, SQL)) Then rs.MoveFirst

Screen.MousePointer = vbArrowHourglass
SQL$ = wInicioTrans
cn.Execute SQL$
Panelprogress.Caption = "Calculando Provision de Vacaciones"

Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rs.RecordCount
Barra.Value = 0

Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   
   raAnt = "": rVaca = "": VacaPag = 0: fing = "": raAct = "": rAcum = ""
   fing = Format(rs!fingreso, "dd/mm/yyyy")
   
   If Mid(fing, 7, 4) = Mid(fecproc, 7, 4) Then
      raAct = Format(Val(Mid(fecproc, 7, 4)) - Mid(fing, 7, 4), "0000") & "."
      raAct = raAct & Format(Val(Mid(fecproc, 4, 2)) - Mid(fing, 4, 2), "00") & "."
      raAct = raAct & Format(Val(Mid(fecproc, 1, 2)) - Mid(fing, 1, 2), "00")
   Else
      raAct = Format(Val(Mid(fecproc, 7, 4)) - Mid(mfecanoant, 7, 4), "0000") & "."
      raAct = raAct & Format(Val(Mid(fecproc, 4, 2)) - Mid(mfecanoant, 4, 2), "00") & "."
      If Val(Mid(fecproc, 1, 2)) - Mid(mfecanoant, 1, 2) < 0 Then
         raAct = raAct & "00"
      Else
         raAct = raAct & Format(Val(Mid(fecproc, 1, 2)) - Mid(mfecanoant, 1, 2), "00")
      End If
   End If
   
   If Mid(raAct, 6, 1) = "-" Then
      raAct = Format(Val(Mid(raAct, 1, 4) - 1), "0000") & "." & Format(Val(Mid(raAct, 6, 3)) + 12, "00") & "." & Right(raAct, 2)
   End If
   
   'BUSCA EL RECORDACU DEL AÑO PASADO
   SQL = "select recordacu from plaprovvaca where cia='" & rs!cia & "' and placod='" & Trim(rs!PLACOD) & "' and year(fechaproceso)=" & Val(Mid(mfecanoant, 7, 4)) & " " _
       & "and month(fechaproceso)=12 and status<>'*'"
       
   'SQL = "SELECT RECOANT AS recordacu FROM VACACIO$ WHERE PLACOD='" & Trim(rs!PLACOD) & "'"
   
   raAnt = "0000.00.00"
   If (fAbrRst(rs2, SQL)) Then
       If Not IsNull(rs2(0)) Then raAnt = rs2(0)
   End If
   
   rs2.Close
     
   SQL = "select h.fechaproceso,h.placod,p.fingreso from plahistorico h,planillas p where h.cia='" & rs!cia & "' and h.proceso='02' and h.placod='" & Trim(rs!PLACOD) & "' " _
       & "and year(fechaproceso)=" & Val(Mid(fecproc, 7, 4)) & " and month(fechaproceso)<=" & Cmbmes.ListIndex + 1 & " and h.status<>'*' " _
       & "and p.cia=h.cia and p.placod=h.placod and p.status<>'*' and h.fechaproceso>p.fingreso"
      
   If (fAbrRst(rs2, SQL)) Then rs2.MoveFirst
   cont1 = 0: VacaPag = 0
   
   Do While Not rs2.EOF
      If Month(rs2(0)) = Val(Mid(fecproc, 4, 2)) Then VacaPag = VacaPag + 1
      cont1 = cont1 + 1
      rs2.MoveNext
   Loop
    
   rVaca = Format(cont1, "0000") & ".00.00"
   rs2.Close
   
'   rAcum = Format(Val(Mid(raAnt, 1, 4)) - Val(Mid(rVaca, 1, 4)) + Val(Mid(raAct, 1, 4)), "0000") & "."
'   rAcum = rAcum & Format(Val(Mid(raAnt, 6, 2)) - Val(Mid(rVaca, 6, 2)) + Val(Mid(raAct, 6, 2)), "00") & "."
'   rAcum = rAcum & Format(Val(Mid(raAnt, 9, 2)) - Val(Mid(rVaca, 9, 2)) + Val(Mid(raAct, 9, 2)), "00")
'
'   If Trim(RS("PLACOD")) = "O5327" Then
'       Debug.Print "HOLA"
'   End If

   rAcum = FECHA_ACUMULADO(rVaca, raAnt, raAct)
      
   SQL$ = "select concepto,tipo,factor_horas,sum(importe) as base from plaremunbase a where cia='" & wcia & "' " _
        & "and placod='" & Trim(rs!PLACOD) & "' and concepto<>'03' and status<>'*' group by concepto,factor_horas," & _
        "A.TIPO order by concepto"
        
   If (fAbrRst(rs2, SQL$)) Then rs2.MoveFirst
   mtoting = 0: mCadIng = "": cont1 = 1
   
   Do While Not rs2.EOF
      mBase = 0
      For i = 1 To 50
          If i = Val(rs2(0)) Then
             If Trim(VTipo) = "01" Then
                If Trim(rs2!tipo) = "04" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             Else
                If Trim(rs2!tipo) = "01" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             End If
             'mCadIng = mCadIng & "" & mBase & "" & ","
             mtoting = mtoting + mBase
             PLAS(cont1) = mBase
             cont1 = cont1 + 1
             Exit For
          ElseIf i > cont1 Then
             mBase = 0
             'mCadIng = mCadIng & "" & mBase & "" & ","
             PLAS(cont1) = mBase
             cont1 = cont1 + 1
          End If
      Next
      rs2.MoveNext
   Loop
    
   'TOTHXS = (RS("i10") + RS("i21") + RS("i24") + RS("I25") + RS("i11")) / 6 / 30
   
   mFactProm = 0
   
   'Promedios
   SQL = "select codinterno,factor from platasaanexo where cia='" & _
   Trim(rs!cia) & "' and modulo='01' and tipomovimiento='" & VTipo & "' and status<>'*'"
   
   If (fAbrRst(rs2, SQL$)) Then rs2.MoveFirst: mFactProm = rs2(1)
   mCadProm = ""
   nFields = 0
   Do While Not rs2.EOF
      mCadProm = mCadProm & "sum(i" & rs2(0) & " ) as i" & rs2(0) & ","
      nFields = nFields + 1
      rs2.MoveNext
   Loop
   rs2.Close
   fpromed = ""
   
   If Trim(mCadProm) <> "" Then
      mCadProm = Mid(mCadProm, 1, Len(Trim(mCadProm)) - 1)
'      mFactProm = Mid(fecproc, 4, 2) - 5
'      mFactProm = DateAdd("m", -5, fecproc)
      fpromed = Fecha_Promedios(mFactProm, fecproc)
      
      If Val(Mid(fecproc, 4, 2)) = 1 Then
         mfectope = "31/12/" & Format(Val(Mid(fecproc, 7, 4)) - 1, "0000")
      Else
         dia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
         mfectope = Format(dia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Mid(fecproc, 7, 4)
      End If
      
      SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      SQL = SQL & " select " & mCadProm & ",sum(I10) as i10,sum(I11) as i11,sum(I21) as i21,sum(I24) as i24,sum(I25) as i25 " & _
      "from plahistorico where cia='" & rs!cia & "' and  placod='" & _
      Trim(rs!PLACOD) & "' and fechaproceso BETWEEN '" & _
      Format(fpromed, FormatFecha) + FormatTimei & "' AND '" & Format(mfectope, FormatFecha) + FormatTimef & "' " _
      & "and proceso='01' and status<>'*'"
      
'      If Trim(rs("PLACOD")) = "O1009" Then
'         Debug.Print "SEGA"
'      End If
          
     If (fAbrRst(rs2, SQL$)) Then rs2.MoveFirst
     mCadProm = ""
        For i = 1 To 50
            mBase = 0
            For j = 0 To nFields - 1
                mBase = 0
                If Format(i, "00") = Mid(rs2(j).Name, 2, 2) Then
                   If IsNull(rs2(j)) Then mBase = 0 Else mBase = rs2(j) / mFactProm
                   Exit For
                End If
            Next j
            mtoting = mtoting + mBase
            mCadProm = mCadProm & "" & mBase & "" & ","
        Next i
        
        
        PLAS(10) = IIf(IsNull(rs2("I10")), 0, rs2("I10"))
        PLAS(21) = IIf(IsNull(rs2("i21")), 0, rs2("i21"))
        PLAS(24) = IIf(IsNull(rs2("I24")), 0, rs2("I24"))
        PLAS(25) = IIf(IsNull(rs2("I25")), 0, rs2("I25"))
        PLAS(11) = IIf(IsNull(rs2("I11")), 0, rs2("I11"))
   
        mCadIng = ""
        For i = 1 To 50
            mCadIng = mCadIng & PLAS(i) & ","
        Next
   
        'mCadIng = Left(mCadIng, Len(mCadIng) - 1)
   End If
   If rs2.State = 1 Then rs2.Close
   
   
   If Val(Mid(fecproc, 4, 2)) - 1 <= 0 Then
      mfectope = "31/01/" & Val(Mid(fecproc, 7, 4) - 1)
   Else
      mfectope = Mid(fecproc, 1, 2) & "/" & Format(Val(Mid(fecproc, 4, 2)) - 1, "00") & "/" & Mid(mfectope, 7, 4)
   End If
  
   SQL = "select provtotal from plaprovvaca where cia='" & rs!cia & _
   "' and placod='" & Trim(rs!PLACOD) & "' and year(fechaproceso)=" & Val(Mid(mfectope, 7, 4)) & " " & _
   "and month(fechaproceso)=" & Val(Mid(mfectope, 4, 2)) & " and status<>'*'"
   
   SQL = "SELECT TOTPER2 FROM VACACIO WHERE PLACOD='" & Trim(rs!PLACOD) & "'"
'   Debug.Print mfectope
   If (fAbrRst(rs2, SQL)) Then mProvAnt = rs2(0) Else mProvAnt = 0
   rs2.Close

   mVacPorPagar = 0
   If VTipo = "01" Then
      mVacPagada = mtoting * VacaPag
      If Val(Mid(raAct, 1, 4)) <> 0 Then mVacPorPagar = (Val(Mid(raAct, 1, 4)) * mtoting)
      If Val(Mid(raAct, 6, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 6, 2)) * mtoting / 12, 2)
      If Val(Mid(raAct, 9, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 9, 2)) * mtoting / 365, 2)
   Else
      mVacPagada = (mtoting * 30) * VacaPag
      If Val(Mid(raAct, 1, 4)) <> 0 Then mVacPorPagar = (Val(Mid(raAct, 1, 4)) * 30 * mtoting)
      If Val(Mid(raAct, 6, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 6, 2)) * 30 * mtoting / 12, 2)
      If Val(Mid(raAct, 9, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 9, 2)) * 30 * mtoting / 365, 2)
   End If
   
   mProvMes = mVacPorPagar - (mProvAnt - mVacPagada)
   
   If mProvMes <> 0 Then
      SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      SQL = SQL & "insert into plaprovvaca values('" & Trim(rs!cia) & "','" & Trim(rs!PLACOD) & "','" & Trim(rs!tipotrabajador) & "', " _
      & "'" & raAnt & "','" & rVaca & "','" & raAct & "','" & rAcum & "'," & mCadIng & mCadProm
      SQL = SQL & "" & mtoting & "," & mProvAnt & "," & mVacPagada & "," & _
      mVacPorPagar & "," & mProvMes & ",'" & Format(fecproc, FormatFecha) & "'," _
      & FechaSys & ",'" & wuser & "','" & Trim(rs!Area) & "','')"
'      Debug.Print SQL
      
      cn.Execute SQL$
   End If
   
   rs.MoveNext
Loop

SQL$ = wFinTrans
cn.Execute SQL$
Panelprogress.Visible = False
Carga_Prov_Vaca
Screen.MousePointer = vbDefault
End Sub

Private Sub SpinButton1_SpinDown()
If Trim(Txtano.Text) = "" Then
   Txtano.Text = Format(Year(Date), "0000")
Else
   Txtano.Text = Txtano.Text - 1
End If
End Sub

Private Sub SpinButton1_SpinUp()
If Trim(Txtano.Text) = "" Then
   Txtano.Text = Format(Year(Date), "0000")
Else
   Txtano.Text = Txtano.Text + 1
End If
End Sub

Private Sub Txtano_Change()
Procesa_Consultas
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Carga_Prov_Vaca()
Dim mcad As String
If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
If VTipo = "99" Or VTipo = "" Then mcad = "" Else mcad = " and a.tipotrab='" & VTipo & "' "

SQL = nombre()
SQL = SQL & "a.placod,a.provmes " _
& "from plaprovvaca a,planillas b " _
& "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*'" & mcad & "order by nombre"

cn.CursorLocation = adUseClient
Set Adocabeza.Recordset = cn.Execute(SQL$, 64)
If Adocabeza.Recordset.RecordCount > 0 Then
   Adocabeza.Recordset.MoveFirst
   Command4.Enabled = False
   SQL = "select SUM(provmes) from plaprovvaca a " _
       & "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*'" & mcad
   
   If (fAbrRst(rs, SQL)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   Lbltotal.Caption = "0.00"
End If
Dgrdcabeza.Refresh
End Sub
Public Sub Elimina_Prov_Vaca()
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If MsgBox("Desea Eliminar Provision ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    SQL = wInicioTrans
    cn.Execute SQL
    
    SQL = "update plaprovvaca set status='*' where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab='" & VTipo & "' and status<>'*'"
    cn.Execute SQL
    
    SQL = wFinTrans
    cn.Execute SQL
    
    Carga_Prov_Vaca
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub Reporte_Provision_Vaca()
Dim marea As String
Dim wciamae As String
Dim mcad As String
Dim mCadBas As String
Dim mCadProm As String
Dim mFieldBas As Integer
Dim mFieldProm As Integer
Dim MField As Integer
Dim contbase As Integer
Dim mText As String
Dim msum As Integer
Dim sumparc As Integer
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency
Dim con As String, CON1 As String
Dim i As Integer
Dim X As Integer
Dim TOTHXS As Variant, TOTASGFAM As Variant
Dim RX As New ADODB.Recordset


If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If Adocabeza.Recordset.RecordCount <= 0 Then Exit Sub

Screen.MousePointer = vbArrowHourglass
mcad = ""
For i = 1 To 50
    mcad = mcad & "sum(i" & Format(i, "00") & "),"
Next

For i = 1 To 50
    mcad = mcad & "sum(p" & Format(i, "00") & "),"
Next

If Trim(mcad) <> "" Then
   mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
End If

SQL = "select " & mcad & " from plaprovvaca " _
    & "where cia='" & wcia & "' and year(fechaproceso)='" & Val(Txtano.Text) & "' and month(fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and tipotrab='" & VTipo & "' and status<>'*'"
    
Panelprogress.Caption = "Generando Reporte de Provision de Vacaciones"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Value = 0
If (fAbrRst(rs, SQL)) Then
   mCadBas = ""
   mCadProm = ""
   Barra.Max = 99
   For i = 0 To 99
       Barra.Value = i
       If rs(i) <> 0 Then
          If i <= 49 Then
             mCadBas = mCadBas & Format(i + 1, "00")
          Else
             mCadProm = mCadProm & Format(i - 50 + 1, "00")
          End If
       End If
   Next
End If

If Trim(mCadBas) = "" Then mFieldBas = 0 Else mFieldBas = Len(Trim(mCadBas))
If Trim(mCadProm) = "" Then mFieldProm = 0 Else mFieldProm = Len(Trim(mCadProm))
marea = ""
wciamae = Determina_Maestro("01044")

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 3
xlSheet.Range("B:B").ColumnWidth = 7
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("D:G").ColumnWidth = 12
xlSheet.Range("D:U").HorizontalAlignment = xlCenter
xlSheet.Range("H:AZ").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 2).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 2).Font.Size = 12
xlSheet.Cells(1, 2).Font.Bold = True

xlSheet.Cells(3, 2).Value = "REPORTE DE PROVISION DE VACACIONES " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(3, 2).Font.Size = 11
xlSheet.Cells(3, 2).Font.Bold = True

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 21)).Font.Bold = True
xlSheet.Cells(6, 2).Value = "Codigo"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 2)).Merge
xlSheet.Cells(6, 2).VerticalAlignment = xlCenter

xlSheet.Cells(6, 3).Value = "Nombre Trabajador"
xlSheet.Range(xlSheet.Cells(6, 3), xlSheet.Cells(7, 3)).Merge
xlSheet.Cells(6, 3).VerticalAlignment = xlCenter

xlSheet.Cells(6, 4).Value = "Record Año"
xlSheet.Cells(7, 4).Value = "Año Mes Dia"

xlSheet.Cells(6, 5).Value = "Vacaciones"
xlSheet.Cells(7, 5).Value = "Tomadas"

xlSheet.Cells(6, 6).Value = "Record Año"
xlSheet.Cells(7, 6).Value = "Actual"

xlSheet.Cells(6, 7).Value = "Record Perdi."
xlSheet.Cells(7, 7).Value = "Año Mes Dia"

xlSheet.Cells(6, 8).Value = "Record Acum."
xlSheet.Cells(7, 8).Value = "Año Mes Dia"

xlSheet.Cells(6, 9) = "Jornal"
xlSheet.Cells(7, 9) = "Basico"
xlSheet.Range(xlSheet.Cells(6, 9), xlSheet.Cells(7, 9)).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 10) = "Promedio"
xlSheet.Cells(7, 10) = "H. Extras"

xlSheet.Cells(6, 11) = "Asignac."
xlSheet.Cells(7, 11) = "familiar"

xlSheet.Cells(6, 12) = "Promedio"
xlSheet.Cells(7, 12) = "Produccion"

xlSheet.Cells(6, 13) = "Bonif."
xlSheet.Cells(7, 13) = "T. Serv."

xlSheet.Cells(6, 14) = "Bonif."
xlSheet.Cells(7, 14) = "Volunt."
xlSheet.Range(xlSheet.Cells(6, 14), xlSheet.Cells(7, 14)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 15) = "AFP"
xlSheet.Range(xlSheet.Cells(6, 15), xlSheet.Cells(7, 15)).Merge
xlSheet.Range(xlSheet.Cells(6, 15), xlSheet.Cells(7, 15)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 16) = "Remun."
xlSheet.Cells(7, 16) = "Vacac."

nCol = 17
contbase = 15

xlSheet.Cells(7, nCol).Value = "Anterior"
xlSheet.Cells(6, nCol).Value = "Mes"
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(7, nCol).Value = "Pagada"
xlSheet.Cells(6, nCol).Value = "Vaca."
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(7, nCol).Value = "Perdidas"
xlSheet.Cells(6, nCol).Value = "Vaca."
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(7, nCol).Value = "Por Pagar"
xlSheet.Cells(6, nCol).Value = "Vaca."
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(6, nCol).Value = "Prov."
xlSheet.Cells(7, nCol).Value = "Del Mes"
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous


xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, nCol)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, nCol)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 21)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 21)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 21)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 21)).Borders(xlInsideVertical).LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, nCol)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 3)).Borders.LineStyle = xlContinuous

nFil = 8
sumparc = nFil
SQL = nombre()
SQL = SQL & "a.* from plaprovvaca a,planillas b " & _
 "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.tipotrab='" & VTipo & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by b.area,nombre"

If (fAbrRst(rs, SQL)) Then rs.MoveFirst: marea = Trim(rs!Area)
If Trim(marea) <> "" Then
   SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & marea & "'"
   SQL$ = SQL$ & wciamae
   If (fAbrRst(rs2, SQL)) Then xlSheet.Cells(nFil, 2).Value = rs2!descrip
   xlSheet.Cells(nFil, 2).Font.Bold = True
   nFil = nFil + 2
   sumparc = sumparc + 2
   rs2.Close
End If

tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0

Barra.Max = rs.RecordCount
Barra.Value = 0

Do While Not rs.EOF

   Barra.Value = rs.AbsolutePosition
   If Trim(rs!Area) <> Trim(marea) Then
      msum = (sumparc - 2) * -1
      nFil = nFil + 1
      X = 1
      For i = nCol - 5 To nCol - 1
         xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         Select Case X
                Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, i).Value
                Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, i).Value
                Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, i).Value
                Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, i).Value
                Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, i).Value
         End Select
         X = X + 1
      Next i
      sumparc = 0
      
      nFil = nFil + 2
      sumparc = sumparc + 3
      marea = rs!Area
      SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & marea & "'"
      SQL$ = SQL$ & wciamae
      If (fAbrRst(rs2, SQL)) Then xlSheet.Cells(nFil, 2).Value = rs2!descrip
      xlSheet.Cells(nFil, 2).Font.Bold = True
      nFil = nFil + 2
      sumparc = sumparc + 2
      rs2.Close
   End If
   
   xlSheet.Cells(nFil, 2).Value = Trim(rs!PLACOD)
   xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre)
   xlSheet.Cells(nFil, 4).Value = RTrim(rs("RECORDANOANT"))
   xlSheet.Cells(nFil, 5).Value = RTrim(rs("RECORDVACA"))
   xlSheet.Cells(nFil, 6).Value = RTrim(rs("RECORDANOACT"))
   
   'RECORD ACUMULADO
   xlSheet.Cells(nFil, 8).Value = rs("RECORDACU")
   
   'jornalbasico
   xlSheet.Cells(nFil, 9) = rs("i01")
   
   TOTHXS = (rs("i10") + rs("i21") + rs("i24") + rs("I25") + rs("i11"))
'   Debug.Print rs("i10")
'   Debug.Print rs("i21")
'   Debug.Print rs("i24")
'   Debug.Print rs("I25")
'   Debug.Print rs("i11")
   
   TOTHXS = (rs("i10") + rs("i21") + rs("i24") + rs("I25") + rs("i11")) / 6 / 30
   
   nCol = 9
   'total horas extras
   xlSheet.Cells(nFil, nCol + 1) = TOTHXS
   
   
   'promedio produccion
   xlSheet.Cells(nFil, nCol + 3) = rs("I18") / 180
   
   con = "SELECT * FROM PLAREMUNBASE WHERE PLACOD='" & Trim(rs("PLACOD")) & "'"
   
   RX.Open con, cn, adOpenStatic, adLockReadOnly
   
   RX.Filter = "CONCEPTO='02'"
   'asignacion familiar
   TOTASGFAM = (IIf(RX.RecordCount = 0, "0.00", RX("importe")) * 4 / 30)
   xlSheet.Cells(nFil, nCol + 2) = TOTASGFAM
   
   
   
   RX.Filter = "CONCEPTO='04'"
   'bonif. t. servicio
   xlSheet.Cells(nFil, nCol + 4) = IIf(RX.RecordCount = 0, "0.00", RX("IMPORTE"))
   
   RX.Filter = "CONCEPTO='26'"
   'bonif. volunt.
   xlSheet.Cells(nFil, nCol + 5) = IIf(RX.RecordCount = 0, "0.00", RX("IMPORTE"))
   
   RX.Filter = "CONCEPTO='06'"
   'afp
   xlSheet.Cells(nFil, nCol + 6) = IIf(RX.RecordCount = 0, "0.00", RX("IMPORTE"))
   
   RX.Close
   
   'REMUNERACION VACACIONAL
   xlSheet.Cells(nFil, nCol + 7) = rs("I06") + rs("I04") + rs("I18") + rs("i01") + TOTHXS + TOTASGFAM
   
   'MES ANTERIOR
   xlSheet.Cells(nFil, nCol + 8) = rs("PROVMESANT")
   
   'VACACIONES PAGADAS
   xlSheet.Cells(nFil, nCol + 9) = rs("PROVPAGADAS")
   
   'VACACIONES PERDIDAS
   xlSheet.Cells(nFil, nCol + 10) = ""
   
   CON1 = xlSheet.Cells(nFil, nCol + 7)
   con = Trim(rs("RECORDACU"))
   tot1 = CON1 * Mid(con, 1, InStr(con, ".")) * 30
   tot2 = CON1 * Mid(con, InStr(con, ".") + 1, Len(con) - InStrRev(con, ".")) * 2.25
   tot3 = (CON1 * 30) / 12 / 30 * Right(con, Len(con) - InStrRev(con, "."))
   
   'VACACIONES POR PAGAR
   xlSheet.Cells(nFil, nCol + 11) = tot1 + tot2 + tot3
   
   'PROV. DEL MES
   xlSheet.Cells(nFil, nCol + 12) = (tot1 + tot2 + tot3) - CON1
   
   nCol = 30
'   If mFieldBas > 0 Then
'      For I = 1 To mFieldBas Step 2
'          MField = Val(Mid(mCadBas, I, 2))
'          xlSheet.Cells(nFil, nCol).Value = RS(MField + 7)
'          nCol = nCol + 1
'      Next
'   End If
   
'   If mFieldProm > 0 Then
'      For I = 1 To mFieldProm Step 2
'          MField = Val(Mid(mCadProm, I, 2))
'          xlSheet.Cells(nFil, nCol).Value = RS(MField + 50 + 7)
'          nCol = nCol + 1
'      Next
'   End If
'
'   xlSheet.Cells(nFil, nCol).Value = RS!remtotal: nCol = nCol + 1
'   xlSheet.Cells(nFil, nCol).Value = RS!provmesant: nCol = nCol + 1
'   xlSheet.Cells(nFil, nCol).Value = RS!provpagadas: nCol = nCol + 1
'   xlSheet.Cells(nFil, nCol).Value = RS!provtotal: nCol = nCol + 1
'   xlSheet.Cells(nFil, nCol).Value = RS!provmes: nCol = nCol + 1
   nFil = nFil + 1
   'sumparc = sumparc + 1
   rs.MoveNext
Loop

Panelprogress.Visible = False

msum = (sumparc - 2) * -1
nFil = nFil + 1
X = 1
For i = nCol - 5 To nCol - 1
    xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
    Select Case X
           Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, i).Value
           Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, i).Value
           Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, i).Value
           Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, i).Value
           Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, i).Value
    End Select
    X = X + 1
Next i
sumparc = 0

nFil = nFil + 1

msum = (nFil) * -1
nFil = nFil + 1
For i = 8 To nCol - 6
   xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next i

X = 1
For i = nCol - 5 To nCol
    Select Case X
           Case Is = 1: xlSheet.Cells(nFil, i).Value = tot1
           Case Is = 2: xlSheet.Cells(nFil, i).Value = tot2
           Case Is = 3: xlSheet.Cells(nFil, i).Value = tot3
           Case Is = 4: xlSheet.Cells(nFil, i).Value = tot4
           Case Is = 5: xlSheet.Cells(nFil, i).Value = tot5
    End Select
    
    X = X + 1
Next

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE VACACIONES"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0
End Sub

Public Sub Provisiones(tipo As String)
Lbltipo.Caption = tipo
If tipo = "V" Then Me.Caption = "Provision de Vacacion"
If tipo = "G" Then Me.Caption = "Provision de Gratificacion"
If tipo = "D" Then
   Me.Caption = "Vacaciones Devengadas"
   Cmbmes.ListIndex = 0
   Cmbmes.Enabled = False
   Label3.Caption = "Total Neto"
   Frame3.BackColor = &HFF&
   Dgrdcabeza.Columns(1).Caption = "Tot. Ing."
End If
End Sub
Public Sub Procesa_Devengadas()
Dim mano As Integer
Dim mmes As Integer
Dim mcad As String
mano = Val(Txtano.Text)
mmes = Cmbmes.ListIndex + 1

If VTipo = "99" Or VTipo = "" Then mcad = "" Else mcad = " and b.tipotrabajador='" & VTipo & "' "

SQL$ = nombre()
SQL$ = SQL$ & "a.placod,a.totaling AS provmes " _
     & "from plahistorico a,planillas b " _
     & "where a.cia='" & wcia & "' and a.proceso='02' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " & mcad
SQL = SQL & "and a.status='D' and a.placod=b.placod and a.cia=b.cia and b.status<>'*'"
cn.CursorLocation = adUseClient

Set Adocabeza.Recordset = cn.Execute(SQL$, 64)
If Adocabeza.Recordset.RecordCount > 0 Then
   Adocabeza.Recordset.MoveFirst
   Dgrdcabeza.Enabled = True
Else
   Dgrdcabeza.Enabled = False
End If
Dgrdcabeza.Refresh
Screen.MousePointer = vbDefault
End Sub
Private Sub Procesa_Consultas()
If Lbltipo.Caption = "V" Then Carga_Prov_Vaca
If Lbltipo.Caption = "D" Then Procesa_Devengadas
If Lbltipo.Caption = "G" Then Carga_Prov_Grati
End Sub
Private Sub Carga_Prov_Grati()
Dim mcad As String

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
If VTipo = "99" Or VTipo = "" Then mcad = "" Else mcad = " and a.tipotrab='" & VTipo & "' "
SQL = nombre()
SQL = SQL & "a.placod,a.gratmes " _
& "from plaprovgrati a,planillas b " _
& "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*'" & mcad & "order by nombre"

cn.CursorLocation = adUseClient
Set Adocabeza.Recordset = cn.Execute(SQL$, 64)
If Adocabeza.Recordset.RecordCount > 0 Then
   Adocabeza.Recordset.MoveFirst
   Command4.Enabled = False
   SQL = "select SUM(gratmes) from plaprovgrati a " _
       & "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*'" & mcad
   
   If (fAbrRst(rs, SQL)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   Lbltotal.Caption = "0.00"
End If
Dgrdcabeza.Refresh
End Sub
Private Sub Calcula_Provision_Grati()
Dim dia As Integer
Dim fecproc As String
Dim mtoting As Currency
Dim mBase As Currency
Dim mCadIng As String
Dim mCadProm As String
Dim cont1 As Integer
Dim fpromed As String
Dim mFactProm As Integer
Dim nFields As Integer
Dim mfectope As String
Dim i As Integer
Dim j As Integer
Dim mGratAnt As Currency
Dim mGratMes As Currency
Dim mProvMes As Currency
Dim mMeses As Integer


dia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
fecproc = Format(dia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text

SQL = "select * from planillas where cia='" & wcia & "' and tipotrabajador='" & VTipo & "' and fcese is null and status<>'*' order by placod"
If (fAbrRst(rs, SQL)) Then rs.MoveFirst

Screen.MousePointer = vbArrowHourglass

Panelprogress.Caption = "Calculando Provision de Gratificaciones"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rs.RecordCount
Barra.Value = 0
SQL$ = wInicioTrans
cn.Execute SQL$

Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   mGratAnt = 0
   'Numero de Meses
   If Cmbmes.ListIndex + 1 = 1 Or Cmbmes.ListIndex + 1 = 7 Then
      mGratAnt = 0: mMeses = 0
      If Year(rs!fingreso) < Val(Txtano.Text) Then
         mMeses = 1
      Else
         If Month(rs!fingreso) < Cmbmes.ListIndex + 1 Then
            mMeses = 1
         Else
            If Day(rs!fingreso) = 1 Then mMeses = 1
         End If
      End If
   Else
      SQL = "select gratmes from plaprovgrati where cia='" & rs!cia & "' and month(fechaproceso)=" & Cmbmes.ListIndex & " and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, SQL)) Then
         mGratAnt = rs2(0): mMeses = rs2(1) + 1:
      Else
        mGratAnt = 0
        If Year(rs!fingreso) < Val(Txtano.Text) Then
           mMeses = 1
           If Month(rs!fingreso) < Cmbmes.ListIndex + 1 Then
              mMeses = 1
           Else
             If Day(rs!fingreso) = 1 Then mMeses = 1
           End If
        End If
      End If
   End If
   
   'Remuneraciones
   SQL$ = "select concepto,tipo,factor_horas,sum(importe) as base from plaremunbase a where cia='" & wcia & "' " _
        & "and placod='" & Trim(rs!PLACOD) & "' and concepto<>'03' " & _
        "and status<>'*' group by concepto,factor_horas,tipo order by concepto"
        
   If (fAbrRst(rs2, SQL$)) Then rs2.MoveFirst
   mtoting = 0: mCadIng = "": cont1 = 0
   Do While Not rs2.EOF
      mBase = 0
      For i = 1 To 50
          If i = Val(rs2(0)) Then
             If VTipo = "01" Then
                If rs2!tipo = "04" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             Else
                If rs2!tipo = "01" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             End If
             mCadIng = mCadIng & "" & mBase & "" & ","
             mtoting = mtoting + mBase
             cont1 = cont1 + 1
             Exit For
          ElseIf i > cont1 Then
             mBase = 0
             mCadIng = mCadIng & "" & mBase & "" & ","
             cont1 = cont1 + 1
          End If
      Next
      rs2.MoveNext
   Loop
   
   For i = (cont1 + 1) To 50
       mCadIng = mCadIng & "0,"
   Next
   mFactProm = 0
   '49
   'Promedios
   SQL = "select codinterno,factor from platasaanexo where cia='" & _
   rs!cia & "' and modulo='01' and tipomovimiento='" & _
   VTipo & "' and status<>'*'"
   
   If (fAbrRst(rs2, SQL$)) Then rs2.MoveFirst: mFactProm = rs2(1)
   mCadProm = ""
   nFields = 0
   Do While Not rs2.EOF
      mCadProm = mCadProm & "sum(i" & rs2(0) & " ) as i" & rs2(0) & ","
      nFields = nFields + 1
      rs2.MoveNext
   Loop
   '4=53
   rs2.Close
   fpromed = ""
   If Trim(mCadProm) <> "" Then
      mCadProm = Mid(mCadProm, 1, Len(Trim(mCadProm)) - 1)
      fpromed = Fecha_Promedios(mFactProm, fecproc)
      
      If Val(Mid(fecproc, 4, 2)) = 1 Then
         mfectope = "31/12/" & Format(Val(Mid(fecproc, 7, 4)) - 1, "0000")
      Else
         dia = Ultimo_Dia(Cmbmes.ListIndex, Val(Txtano.Text))
         mfectope = Format(dia, "00") & "/" & Format(Cmbmes.ListIndex, "00") & "/" & Mid(fecproc, 7, 4)
      End If
      
      SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql
      SQL = SQL & " select " & mCadProm & " from plahistorico where cia='" & rs!cia & "' and  placod='" & Trim(rs!PLACOD) & "' and fechaproceso " _
          & "BETWEEN '" & Format(fpromed, FormatFecha) + FormatTimei & "' AND '" & Format(mfectope, FormatFecha) + FormatTimef & "' " _
          & "and proceso='01' and status<>'*'"
          
     If (fAbrRst(rs2, SQL$)) Then rs2.MoveFirst
     mCadProm = ""
        For i = 1 To 50
            mBase = 0
            For j = 0 To nFields - 1
                If Format(i, "00") = Mid(rs2(j).Name, 2, 2) Then
                   If IsNull(rs2(j)) Then mBase = 0 Else mBase = rs2(j) / mFactProm
                   Exit For
                Else
                   mBase = 0
                End If
            Next j
            mtoting = mtoting + mBase
            mCadProm = mCadProm & "" & mBase & "" & ","
        Next i
   End If
   rs2.Close
   '103
   
   If mMeses <> 0 Then
      mGratMes = mMeses * mtoting / 6
      mProvMes = mGratMes - mGratAnt
      If mProvMes <> 0 Then
         SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
         SQL = SQL & "insert into plaprovgrati values('" & Trim(wcia) & _
         "','" & Trim(rs!PLACOD) & "','" & Trim(rs!tipotrabajador) & "'," _
         & "" & mMeses & ",'' ,'' ,'' ," & mCadIng & mCadProm
         SQL = SQL & "" & mtoting & "," & mtoting & "," & mGratAnt & _
         "," & mGratMes & "," & mProvMes & ",'" & Format(fecproc, FormatFecha) & _
         "'," & FechaSys & ",'" & wuser & "','" & Trim(rs!Area) & "','')"
             
         cn.Execute SQL$
         
      End If
   End If
   rs.MoveNext
Loop

SQL$ = wFinTrans
cn.Execute SQL$
Panelprogress.Visible = False
Carga_Prov_Grati
Screen.MousePointer = vbDefault

End Sub
Public Sub Elimina_Prov_Grati()
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If MsgBox("Desea Eliminar Provision ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    SQL = wInicioTrans
    cn.Execute SQL
    
    SQL = "update plaprovgrati set status='*' where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab='" & VTipo & "' and status<>'*'"
    cn.Execute SQL
    
    SQL = wFinTrans
    cn.Execute SQL
    
    Carga_Prov_Grati
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub Reporte_Provision_Grati()
Dim marea As String
Dim wciamae As String
Dim mcad As String
Dim mCadBas As String
Dim mCadProm As String
Dim mFieldBas As Integer
Dim mFieldProm As Integer
Dim MField As Integer
Dim contbase As Integer
Dim mText As String
Dim msum As Integer
Dim sumparc As Integer
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency

Dim i As Integer
Dim X As Integer
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If Adocabeza.Recordset.RecordCount <= 0 Then Exit Sub

Screen.MousePointer = vbArrowHourglass
mcad = ""
For i = 1 To 50
    mcad = mcad & "sum(i" & Format(i, "00") & "),"
Next
For i = 1 To 50
    mcad = mcad & "sum(p" & Format(i, "00") & "),"
Next
If Trim(mcad) <> "" Then
   mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
End If

Panelprogress.Caption = "Generando Reporte de Provision de Vacaciones"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Value = 0
SQL = "select " & mcad & " from plaprovgrati " _
    & "where cia='" & wcia & "' and year(fechaproceso)='" & Val(Txtano.Text) & "' and month(fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and tipotrab='" & VTipo & "' and status<>'*'"
    
If (fAbrRst(rs, SQL)) Then
   mCadBas = ""
   mCadProm = ""
   Barra.Max = 99
   For i = 0 To 99
       Barra.Value = i
       If rs(i) <> 0 Then
          If i <= 49 Then
             mCadBas = mCadBas & Format(i + 1, "00")
          Else
             mCadProm = mCadProm & Format(i - 50 + 1, "00")
          End If
       End If
   Next
End If
If Trim(mCadBas) = "" Then mFieldBas = 0 Else mFieldBas = Len(Trim(mCadBas))
If Trim(mCadProm) = "" Then mFieldProm = 0 Else mFieldProm = Len(Trim(mCadProm))
marea = ""
wciamae = Determina_Maestro("01044")

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 3
xlSheet.Range("B:B").ColumnWidth = 7
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("D:G").ColumnWidth = 12
xlSheet.Range("D:AZ").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 2).Value = Cmbcia.Text
xlSheet.Cells(1, 2).Font.Size = 12
xlSheet.Cells(1, 2).Font.Bold = True

xlSheet.Cells(3, 2).Value = "REPORTE DE PROVISION DE VACACIONES " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(3, 2).Font.Size = 11
xlSheet.Cells(3, 2).Font.Bold = True

xlSheet.Cells(6, 2).Value = "Codigo"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 2)).Merge
xlSheet.Cells(6, 2).VerticalAlignment = xlCenter
xlSheet.Cells(6, 3).Value = "Nombre Trabajador"
xlSheet.Range(xlSheet.Cells(6, 3), xlSheet.Cells(7, 3)).Merge
xlSheet.Cells(6, 3).VerticalAlignment = xlCenter

nCol = 4
If mFieldBas > 0 Then
  Barra.Value = 0
  Barra.Max = mFieldBas
  For i = 1 To mFieldBas Step 2
      Barra.Value = i
      MField = Val(Mid(mCadBas, i, 2))
      SQL = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno='" & Format(MField, "00") & "' and status<>'*'"
      If (fAbrRst(rs, SQL)) Then xlSheet.Cells(6, nCol).Value = Trim(rs!descripcion)
      xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Merge
      xlSheet.Cells(6, nCol).HorizontalAlignment = xlCenter
      xlSheet.Cells(6, nCol).VerticalAlignment = xlJustify
      nCol = nCol + 1
  Next
End If
contbase = nCol - 1
If mFieldProm > 0 Then
  Barra.Value = 0
  Barra.Max = mFieldProm
  For i = 1 To mFieldProm Step 2
     Barra.Value = i
      MField = Val(Mid(mCadProm, i, 2))
      xlSheet.Cells(6, nCol).Value = "Promedio"
      SQL = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno='" & Format(MField, "00") & "' and status<>'*'"
      If (fAbrRst(rs, SQL)) Then xlSheet.Cells(7, nCol).Value = Trim(rs!descripcion)
      xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous
      nCol = nCol + 1
  Next
End If

xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Cells(7, nCol).Value = "Grati."
xlSheet.Cells(6, nCol).Value = "Remun.": nCol = nCol + 1

xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Cells(7, nCol).Value = "Anterior"
xlSheet.Cells(6, nCol).Value = "Mes": nCol = nCol + 1

xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Cells(7, nCol).Value = "Por Pagar"
xlSheet.Cells(6, nCol).Value = "Grati.": nCol = nCol + 1

xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Cells(6, nCol).Value = "Prov."
xlSheet.Cells(7, nCol).Value = "Del Mes"

xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, nCol)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, nCol)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, nCol)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 3)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, contbase)).Borders.LineStyle = xlContinuous

nFil = 8
sumparc = nFil
SQL = nombre()
SQL = SQL & "a.* " _
& "from plaprovgrati a,planillas b " _
& "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.tipotrab='" & VTipo & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by b.area,nombre"

If (fAbrRst(rs, SQL)) Then rs.MoveFirst: marea = rs!Area
If marea <> "" Then
   SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & Trim(marea) & "'"
   SQL$ = SQL$ & wciamae
  
   If (fAbrRst(rs2, SQL)) Then xlSheet.Cells(nFil, 2).Value = rs2!descrip
   xlSheet.Cells(nFil, 2).Font.Bold = True
   nFil = nFil + 2
   sumparc = sumparc + 2
   rs2.Close
End If

tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
Barra.Value = 0
Barra.Max = rs.RecordCount
Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   If rs!Area <> marea Then
      msum = (sumparc - 2) * -1
      nFil = nFil + 1
      X = 1
      For i = nCol - 4 To nCol - 1
         xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         Select Case X
                Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, i).Value
                Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, i).Value
                Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, i).Value
                Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, i).Value
                Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, i).Value
         End Select
         X = X + 1
      Next i
      sumparc = 0
      
      nFil = nFil + 2
      sumparc = sumparc + 3
      marea = rs!Area
      SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & marea & "'"
      SQL$ = SQL$ & wciamae
      If (fAbrRst(rs2, SQL)) Then xlSheet.Cells(nFil, 2).Value = rs2!descrip
      xlSheet.Cells(nFil, 2).Font.Bold = True
      nFil = nFil + 2
      sumparc = sumparc + 2
      rs2.Close
   End If
   xlSheet.Cells(nFil, 2).Value = rs!PLACOD
   xlSheet.Cells(nFil, 3).Value = rs!nombre
   nCol = 4
   If mFieldBas > 0 Then
      For i = 1 To mFieldBas Step 2
          MField = Val(Mid(mCadBas, i, 2))
          xlSheet.Cells(nFil, nCol).Value = rs(MField + 4)
          nCol = nCol + 1
      Next
   End If
   
   If mFieldProm > 0 Then
      For i = 1 To mFieldProm Step 2
          MField = Val(Mid(mCadProm, i, 2))
          xlSheet.Cells(nFil, nCol).Value = rs(MField + 50 + 7)
          nCol = nCol + 1
      Next
   End If
   
    'SEGA
   xlSheet.Cells(nFil, nCol).Value = rs!remtotal: nCol = nCol + 1
   xlSheet.Cells(nFil, nCol).Value = rs!gratmesant: nCol = nCol + 1
      xlSheet.Cells(nFil, nCol).Value = rs!gratmes: nCol = nCol + 1
   'xlSheet.Cells(nFil, nCol).Value = RS!provmes
   nCol = nCol + 1
   
   nFil = nFil + 1
   sumparc = sumparc + 1
   rs.MoveNext
Loop
Panelprogress.Visible = False
msum = (sumparc - 2) * -1
nFil = nFil + 1
X = 1
For i = nCol - 4 To nCol - 1
    xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
    Select Case X
           Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, i).Value
           Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, i).Value
           Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, i).Value
           Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, i).Value
           Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, i).Value
    End Select
    X = X + 1
Next i
sumparc = 0

nFil = nFil + 1

msum = (nFil) * -1
nFil = nFil + 1
For i = 4 To nCol - 5
   xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next i

X = 1
For i = nCol - 4 To nCol - 1
    Select Case X
           Case Is = 1: xlSheet.Cells(nFil, i).Value = tot1
           Case Is = 2: xlSheet.Cells(nFil, i).Value = tot2
           Case Is = 3: xlSheet.Cells(nFil, i).Value = tot3
           Case Is = 4: xlSheet.Cells(nFil, i).Value = tot4
           Case Is = 5: xlSheet.Cells(nFil, i).Value = tot5
    End Select
    
    X = X + 1
Next

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE GRATIFICACION"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0
End Sub

Function FECHA_ACUMULADO(fecha As Variant, FEC1 As Variant, FEC2 As Variant) As Variant
            
    If Left(FEC1, 1) = "-" Then FEC1 = Mid(FEC1, 2)
    If Left(FEC2, 1) = "-" Then FEC2 = Mid(FEC2, 2)
    Dim ANNO As Variant, ANNO1 As Variant, ANNO2 As Variant
    Dim DIA1 As Variant, DIA2 As Variant, MES1 As Variant
    Dim MES2 As Variant, mes As Variant, dia As Variant
    
    If fecha = "0000.00.00" Then
        FECHA_ACUMULADO = fecha
        Exit Function
    End If
    ANNO = Val(Left(fecha, 4))
    ANNO1 = Left(FEC1, 1)
    ANNO2 = Val(Left(FEC2, 4))
    DIA1 = Mid(FEC1, 6)
    DIA2 = Val(Mid(FEC2, 9))
    MES1 = Mid(FEC1, 3, 2)
    MES2 = Val(Mid(FEC2, 6, 2))
    
    ANNO = ANNO1 - ANNO
    mes = (Val(MES1) + MES2)
    If (mes) > 12 Then
        mes = (mes) - 12
        ANNO = ANNO + 1
    End If
    
    mes = Format(mes, "00")
    dia = Format(Val(DIA1) + Val(DIA2), "00")
    
    FECHA_ACUMULADO = ANNO & "." & mes & "." & dia
           
End Function



Private Sub PROVISION_VACACIONES()

Dim sSQL As String
Dim MaxRow As Long
Dim rs As ADODB.Recordset
Dim EXTRAS As String, PRODUCCION As String, OTROSPAGOS As String
Dim FechaProc As String
Dim FecFin As String, FecIni As String, fecha1 As String
Dim RcdAñoPasado As String, VacTomadas As String, RcdAñoActual As String, RcdAñoTotal As String, ImpProvVacaAnt As String
Dim RSBUSCAR As ADODB.Recordset
Dim strCodigo As String
Dim MaxCol As Integer
Dim dblFactor As Currency
Dim MaxColInicial As Integer
Dim MaxColFin As Integer
Dim MaxColTemp As Integer
Dim i As Integer
Dim Campo As String

Dim sSQLI As String
Dim sSQLP As String
Dim sCol As Integer

MaxRow = 0

' SETEAMOS LAS FECHAS A TRABAJAR
If Cmbmes.ListIndex + 1 = 12 Then
    FechaProc = "01/01/" & Txtano.Text + 1
Else
    FechaProc = "01/" & Format(Cmbmes.ListIndex + 2, "00") & "/" & Txtano.Text
End If

FecFin = DateAdd("d", -1, CDate(FechaProc))
FecIni = DateAdd("m", -6, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))
'*****************************************************************************

sSQL = "select 'I' + campo from estructura_provisiones where campo not in"
sSQL = sSQL & " (select codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
sSQL = sSQL & " Union All select 'P'+codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16'"

Erase ArrReporte

MaxCol = 6
MaxColInicial = MaxCol
i = MaxCol

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        MaxCol = MaxCol + 1
        rs.MoveNext
    Loop
    'AGREGAMOS LOS CAMPOS QUE FALTAN DE LOS IMPORTES
    MaxCol = MaxCol + 5
    
    rs.MoveFirst
    ReDim Preserve ArrReporte(0 To MaxCol, 0 To MaxRow)
        
    Do While Not rs.EOF
        ArrReporte(i, MaxRow) = rs(0)
        i = i + 1
        rs.MoveNext
    Loop
    MaxRow = MaxRow + 1
End If


'OBTENEMOS LOS CONCEPTOS REMUNERATIVOS A TRABAJAR
'-- HORAS EXTRAS --
sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='1'"

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        EXTRAS = EXTRAS & "P" & Trim(rs!codinterno) & "=SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)),"
        rs.MoveNext
    Loop
    rs.MoveFirst
    EXTRAS = Mid(EXTRAS, 1, Len(Trim(EXTRAS)) - 1)
    rs.Close
Else
    EXTRAS = "EXTRAS=0"
End If

'-- HORAS PRODUCCION --
sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='2'"
If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        PRODUCCION = PRODUCCION & "P" & Trim(rs!codinterno) & "=SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)),"
        rs.MoveNext
    Loop
    rs.MoveFirst
    PRODUCCION = Mid(PRODUCCION, 1, Len(Trim(PRODUCCION)) - 1)
    rs.Close
Else
    PRODUCCION = "PRODUCCION=0"
End If

'-- HORAS OTROS PAGOS --
sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='3'"
If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        OTROSPAGOS = OTROSPAGOS & "P" & Trim(rs!codinterno) & "=SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)),"
        rs.MoveNext
    Loop
    rs.MoveFirst
    OTROSPAGOS = Mid(OTROSPAGOS, 1, Len(Trim(OTROSPAGOS)) - 1)
    rs.Close
Else
    OTROSPAGOS = "OTROSPAGOS=0"
End If

sSQL = " SELECT FACTOR FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='3'"
If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
End If
'******************************************************************************************************************************

sSQL = "SET DATEFORMAT DMY SELECT"
sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
sSQL = sSQL & " fingreso , fcese, b.placod, b.codinterno, b.descripcion, b.importe,a.area, "
sSQL = sSQL & " " & EXTRAS & ", " & PRODUCCION & ", " & OTROSPAGOS
sSQL = sSQL & " FROM planillas a INNER JOIN ("
sSQL = sSQL & " SELECT prb.PLACOD , pc.codinterno, pc.descripcion, prb.importe FROM plaremunbase prb INNER JOIN placonstante pc ON"
sSQL = sSQL & " (pc.cia='" & wcia & "' and pc.tipomovimiento='02' and pc.status!='*' and pc.codinterno=prb.concepto) WHERE"
sSQL = sSQL & " prb.STATUS!='*' ) B ON (b.placod=a.placod) LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod)"
sSQL = sSQL & " WHERE a.status!='*' and LEFT(a.placod,1)='T' and a.tipotrabajador='01' and a.cia='" & wcia & "'"
sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, fingreso, fcese, b.PLACOD, b.codinterno, b.descripcion, b.importe,a.area"

fecha1 = "01/01/" & Txtano.Text
If (fAbrRst(rs, sSQL)) Then
    Set RSBUSCAR = rs.Clone
    
Do While Not rs.EOF
    If strCodigo <> Trim(rs!PLACOD) Then
        ReDim Preserve ArrReporte(0 To MaxCol, 0 To MaxRow)
        strCodigo = Trim(rs!PLACOD)
        ObtenerFechas rs!PLACOD, rs!fingreso, fecha1, FecFin, RcdAñoPasado, VacTomadas, RcdAñoActual, RcdAñoTotal, ImpProvVacaAnt
        
        ArrReporte(COL_CODIGO, MaxRow) = Trim(rs!PLACOD)
        ArrReporte(COL_NOMBRE, MaxRow) = Trim(rs(1))
        ArrReporte(COL_RCDAÑOANTERIOR, MaxRow) = RcdAñoPasado
        ArrReporte(COL_VACACTOMADAS, MaxRow) = VacTomadas
        ArrReporte(COL_RCDAÑOACTUAL, MaxRow) = RcdAñoActual
        ArrReporte(COL_RCDACUMULADO, MaxRow) = RcdAñoTotal
        
        'LLENADO DE LOS CAMPOS SEGUN TABLA ESTRUCTURA_PROVISIONES
        For i = MaxColInicial To MaxCol - 6
            If Left(Trim(ArrReporte(i, 0)), 1) = "I" Then
                RSBUSCAR.Filter = "placod='" & Trim(rs!PLACOD) & "' and codinterno='" & Right(Trim(ArrReporte(i, 0)), 2) & "'"
                If Not RSBUSCAR.EOF Then
                    ArrReporte(i, MaxRow) = RSBUSCAR!importe
                Else
                    ArrReporte(i, MaxRow) = 0
                End If
                RSBUSCAR.Filter = ""
            ElseIf Left(Trim(ArrReporte(i, 0)), 1) = "P" Then
                If rs.Fields(Trim(ArrReporte(i, 0))) > 0 Then
                    ArrReporte(i, MaxRow) = rs.Fields(Trim(ArrReporte(i, 0))) / dblFactor
                Else
                    ArrReporte(i, MaxRow) = 0
                End If
            End If
        Next
        
        'SUMAMOS TODOS LOS CAMPOS DE LOS IMPORTES
        For MaxColTemp = MaxColInicial To MaxCol - 6
            ArrReporte(i, MaxRow) = ArrReporte(i, MaxRow) + ArrReporte(MaxColTemp, MaxRow)
        Next MaxColTemp
        i = i + 1
                        
        'IMPORTE DE LAS VACACIONES ANTERIOR
        ArrReporte(i, MaxRow) = ImpProvVacaAnt
        i = i + 1
        'IMPORTE DE LAS VACACIONES TOMADAS
        ArrReporte(i, MaxRow) = CalculaImpVaca(VacTomadas, ArrReporte(MaxColTemp, MaxRow))
        i = i + 1
        'IMPORTE DE LAS VACACIONES X PAGAR
        ArrReporte(i, MaxRow) = CalculaImpVaca(RcdAñoTotal, ArrReporte(MaxColTemp, MaxRow))
        i = i + 1
        'IMPORTE DE LAS PROVISIONES DEL MES
        ArrReporte(i, MaxRow) = Abs(ArrReporte(MaxColTemp + 3, MaxRow) - ArrReporte(MaxColTemp + 1, MaxRow) - ArrReporte(MaxColTemp + 2, MaxRow))
        i = i + 1
        'AREA DEL PERSONAL
        ArrReporte(i, MaxRow) = Trim(rs!Area)
        
        MaxRow = MaxRow + 1
    
        'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
        sSQLI = ""
        For i = 1 To 50
            Campo = "I" & Format(i, "00")
            sCol = BuscaColumna(Campo, MaxCol)
            If sCol > 0 Then
                sSQLI = sSQLI & ArrReporte(sCol, MaxRow - 1) & ","
            Else
                sSQLI = sSQLI & "0,"
            End If
        Next
        
        'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
        sSQLP = ""
        For i = 1 To 50
            Campo = "P" & Format(i, "00")
            sCol = BuscaColumna(Campo, MaxCol)
            If sCol > 0 Then
                sSQLP = sSQLP & ArrReporte(sCol, MaxRow - 1) & ","
            Else
                sSQLP = sSQLP & "0,"
            End If
        Next
        
        sSQL = "INSERT plaprovvaca VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MaxRow - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','" & ArrReporte(COL_RCDAÑOANTERIOR, MaxRow - 1) & "',"
        sSQL = sSQL & "'" & ArrReporte(COL_VACACTOMADAS, MaxRow - 1) & "','" & ArrReporte(COL_RCDAÑOACTUAL, MaxRow - 1) & "','" & ArrReporte(COL_RCDACUMULADO, MaxRow - 1) & "'," & sSQLI & sSQLP & ArrReporte(MaxColTemp, MaxRow - 1) & ","
        sSQL = sSQL & ArrReporte(MaxColTemp + 1, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 2, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 3, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 4, MaxRow - 1) & ",'" & FecFin & "',GETDATE(),'" & wuser & "','" & ArrReporte(MaxColTemp + 5, MaxRow - 1) & "',' ')"
        
        cn.Execute (sSQL)
    End If
    rs.MoveNext
Loop
    
End If

End Sub
Private Function BuscaColumna(ByVal pCampo As String, ByVal pMaxcol As Integer) As Integer
Dim iRow As Integer
BuscaColumna = 0
For iRow = 3 To pMaxcol
    If ArrReporte(iRow, 0) = pCampo Then
        BuscaColumna = iRow
        Exit Function
    End If
Next
End Function


Private Function CalculaImpVaca(ByVal pFecha As String, ByVal pImporteRemun As String) As String
Dim ArrFecha As Variant
Dim i As Integer
Dim CalculaImpVacaTmp As Currency
Dim año As Currency, mes As Currency, dia As Currency

ArrFecha = Split(pFecha, ".")

año = pImporteRemun
mes = pImporteRemun / 12
dia = pImporteRemun / 365

For i = 0 To UBound(ArrFecha)
    Select Case i
    Case Is = 0
        ArrFecha(i) = año * Val(ArrFecha(i))
        CalculaImpVacaTmp = CalculaImpVacaTmp + ArrFecha(i)
    Case Is = 1
        ArrFecha(i) = mes * Val(ArrFecha(i))
        CalculaImpVacaTmp = CalculaImpVacaTmp + ArrFecha(i)
    Case Is = 2
        ArrFecha(i) = dia * Val(ArrFecha(i))
        CalculaImpVacaTmp = CalculaImpVacaTmp + ArrFecha(i)
    End Select
Next i

CalculaImpVaca = Format(CalculaImpVacaTmp, "#0.00")

End Function

Private Sub ObtenerFechas(ByVal pPlacod As String, ByVal pFecIng As String, ByVal pFecIniProc As String, ByVal pFecproc As String, ByRef pRcdAcumuPasado As String, ByRef pVacacTomadas As String, ByRef pRcdActual As String, ByRef pRcdAcumulado As String, ByRef pImpVacaAnterior As String)
Dim sSQL As String
Dim resultado As String
Dim FecInicio As String
Dim FecIngTmp As String
Dim sAÑO As String, sMES As String, sDIA As String


FecIngTmp = pFecIng

sSQL = "SELECT recordacu FROM plaprovvaca WHERE PLACOD='" & Trim(pPlacod) & "' AND STATUS!='*' AND YEAR(fechaproceso)=" & Val(Txtano.Text) - 1 & " AND MONTH(fechaproceso)=12"
If (fAbrRst(rs, sSQL)) Then
    resultado = rs(0)
Else
    resultado = " 0. 0. 0"
End If
pRcdAcumuPasado = resultado

sSQL = "SELECT provtotal FROM plaprovvaca WHERE PLACOD='" & Trim(pPlacod) & "' AND STATUS!='*' AND YEAR(fechaproceso)=" & Year(DateAdd("m", -1, pFecIniProc)) & " AND MONTH(fechaproceso)=" & Month(DateAdd("m", -1, pFecIniProc))
If (fAbrRst(rs, sSQL)) Then
    resultado = rs(0)
Else
    resultado = "0"
End If
pImpVacaAnterior = resultado


sSQL = "SELECT COUNT(*) FROM PLAHISTORICO WHERE placod='" & Trim(pPlacod) & "' and status!='*' and proceso='2' and year(fechaproceso)=" & Year(pFecproc)
If (fAbrRst(rs, sSQL)) Then
    resultado = Space(2 - Len(Trim(rs(0)))) & Trim(rs(0)) & ". 0. 0"
Else
    resultado = " 0. 0. 0"
End If

pVacacTomadas = resultado

If CDate(pFecIng) > CDate(pFecIniProc) Then
    sAÑO = DateDiff("y", CDate(pFecproc), CDate(FecIngTmp))
    FecIngTmp = DateAdd("y", Val(sAÑO), CDate(FecIngTmp))
    
    sMES = DateDiff("m", CDate(pFecproc), CDate(FecIngTmp))
    FecIngTmp = DateAdd("y", Val(sMES), CDate(FecIngTmp))
    
    sDIA = DateDiff("d", CDate(pFecproc), CDate(FecIngTmp))
Else
    sAÑO = "0"
    sMES = Cmbmes.ListIndex + 1
    sDIA = "0"
End If

masdias:
If sDIA >= 30 Then
    sMES = Val(sMES) + 1
    sDIA = Val(sDIA) - 30
    GoTo masdias
End If

masmes:
If sMES >= 12 Then
    sAÑO = Val(sAÑO) + 1
    sMES = Val(sMES) - 12
    GoTo masmes
End If

pRcdActual = Space(2 - Len(Trim(sAÑO))) & sAÑO & "." & Space(2 - Len(Trim(sMES))) & sMES & "." & Space(2 - Len(Trim(sDIA))) & sDIA

sAÑO = "": sMES = "": sDIA = ""

sAÑO = Val(Mid(pRcdAcumuPasado, 1, 2)) - Val(Mid(pVacacTomadas, 1, 2))
sMES = Val(Mid(pRcdAcumuPasado, 4, 2)) - Val(Mid(pVacacTomadas, 4, 2))
sDIA = Val(Mid(pRcdAcumuPasado, 7, 2)) - Val(Mid(pVacacTomadas, 7, 2))

resultado = Space(2 - Len(Trim(sAÑO))) & sAÑO & "." & Space(2 - Len(Trim(sMES))) & sMES & "." & Space(2 - Len(Trim(sDIA))) & sDIA

sAÑO = "": sMES = "": sDIA = ""
sAÑO = Val(Mid(resultado, 1, 2)) + Val(Mid(pRcdActual, 1, 2))
sMES = Val(Mid(resultado, 4, 2)) + Val(Mid(pRcdActual, 4, 2))
sDIA = Val(Mid(resultado, 7, 2)) + Val(Mid(pRcdActual, 7, 2))

masdiasACTUAL:
If sDIA >= 30 Then
    sMES = Val(sMES) + 1
    sDIA = Val(sDIA) - 30
    GoTo masdiasACTUAL
End If

masmesACTUAL:
If sMES >= 12 Then
    sAÑO = Val(sAÑO) + 1
    sMES = Val(sMES) - 12
    GoTo masmesACTUAL
End If

pRcdAcumulado = Space(2 - Len(Trim(sAÑO))) & sAÑO & "." & Space(2 - Len(Trim(sMES))) & sMES & "." & Space(2 - Len(Trim(sDIA))) & sDIA

End Sub

Private Sub PROVICIONES_GRATI()
Dim sSQL As String
Dim MaxRow As Long, MaxCol As Integer, MaxColInicial As Integer
Dim rs As ADODB.Recordset, rsAUX As ADODB.Recordset
Dim CantMes As String, Campo As String
Dim FecIni As String, FecFin As String, FecProceso As String
Dim i As Integer, MaxColTemp As Integer
Dim dblFactor As Currency, CADENA As String
Dim factor_essalud As Currency, totaportes As Currency
Dim sCol As Integer, curfactor As Currency
Dim sSQLI As String, sSQLP As String

Const COL_CODIGO = 0
Const COL_FECING = 1
Const COL_AREA = 2

MaxCol = 2
i = MaxCol + 1
FecProceso = Format(DateAdd("d", -1, Format(DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")

If Cmbmes.ListIndex + 1 < 7 Then
    FecIni = "01/01/" & Txtano.Text
    FecFin = Format(DateAdd("d", -1, "01/07/" & Txtano.Text), "DD/MM/YYYY")
Else
    FecIni = "01/07/" & Txtano.Text
    FecFin = Format(DateAdd("d", -1, CDate("01/01/" & Txtano.Text + 1)), "DD/MM/YYYY")
End If


Erase ArrReporte

sSQL = "SELECT concepto,campo,sn_promedio,0 AS factor,sn_carga FROM estructura_provisiones WHERE tipo='G' and "
sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
sSQL = sSQL & " UNION ALL SELECT concepto,campo,sn_promedio,b.factor,sn_carga FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.codinterno=a.campo) WHERE a.tipo='G'"

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        MaxCol = MaxCol + 1
        rs.MoveNext
    Loop
    rs.MoveFirst
    MaxColTemp = MaxCol + 1
    MaxCol = MaxCol + 5
    
    ReDim Preserve ArrReporte(0 To MaxCol, 0 To MaxRow)
    
    Do While Not rs.EOF
        ArrReporte(i, MaxRow) = Trim(rs!concepto) & rs!Campo
        If CInt(rs!sn_carga) <> 0 Then
            CADENA = CADENA & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=SUM(COALESCE(" & Trim(rs!concepto) & Trim(rs!Campo) & ",0))"
            If CInt(rs!sn_promedio) = -1 And Trim(rs!concepto) = "A" Then
                CADENA = CADENA & "/(SELECT aportacion FROM PLACONSTANTE WHERE tipomovimiento='03' AND cia='" & wcia & "' AND status!='*' AND codinterno='" & rs!Campo & "'),"
            ElseIf CInt(rs!sn_promedio) = -1 And Trim(rs!concepto) = "I" Then
                CADENA = CADENA & "/" & rs!factor & ","
                Else
                    CADENA = CADENA & ","
            End If
        Else
            CADENA = CADENA & "COALESCE((SELECT prb.importe/(factor_horas/" & hORAS_X_DIA & ") FROM plaremunbase prb WHERE placod=a.placod AND status!='*' and concepto='" & Trim(rs!Campo) & "'),0) as '" & Trim(rs!concepto) & Trim(rs!Campo) & "',"
        End If
        
        i = i + 1
        rs.MoveNext
    Loop
    
    CADENA = Mid(CADENA, 1, Len(Trim(CADENA)) - 1)
    rs.Close
End If

sSQL = "SELECT APORTACION FROM PLACONSTANTE WHERE TIPOMOVIMIENTO='03' AND CODINTERNO='01' AND STATUS!='*' AND CIA='" & wcia & "'"
If (fAbrRst(rs, sSQL)) Then
    factor_essalud = rs(0)
    rs.Close
End If

sSQL = "SET DATEFORMAT DMY SELECT"
sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
sSQL = sSQL & " fingreso , fcese, a.area,a.tipotrabajador,"
sSQL = sSQL & CADENA
sSQL = sSQL & " FROM planillas a LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod)"
sSQL = sSQL & " WHERE a.status!='*' and LEFT(a.placod,1)='T' and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "'"
sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"
 
 MaxRow = MaxRow + 1
If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        ReDim Preserve ArrReporte(0 To MaxCol, 0 To MaxRow)
        CantMes = CantidadMesesCalculo(rs!fingreso)
        ArrReporte(COL_CODIGO, MaxRow) = Trim(rs!PLACOD)
        ArrReporte(COL_FECING, MaxRow) = rs!fingreso
        ArrReporte(COL_AREA, MaxRow) = rs!Area
        For i = 6 To rs.Fields.Count - 1
            sCol = BuscaColumna(rs.Fields(i).Name, MaxCol)
            If sCol > 0 Then
                If Trim(rs!tipotrabajador) = "01" Then
                    ArrReporte(sCol, MaxRow) = Round(rs.Fields(i).Value * DIAS_TRABAJO, 2)
                Else
                    ArrReporte(sCol, MaxRow) = Round(rs.Fields(i).Value, 2)
                End If
                If Left(rs.Fields(i).Name, 1) = "I" Then totaportes = totaportes + ArrReporte(sCol, MaxRow)
            End If
        Next
        
        i = MaxColTemp
        ArrReporte(i, 0) = "P01"
        ArrReporte(i, MaxRow) = Round(totaportes * (factor_essalud / 100), 2)
        i = i + 1
        ArrReporte(i, MaxRow) = Round(totaportes + ArrReporte(i - 1, MaxRow), 2)
        totaportes = ArrReporte(i, MaxRow)
        i = i + 1
        
        If Cmbmes.ListIndex + 1 = 1 Or Cmbmes.ListIndex + 1 = 7 Then
            SQL = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
        Else
            SQL = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
        End If
        
        If (fAbrRst(rsAUX, SQL)) Then
            ArrReporte(i, MaxRow) = rsAUX(0)
            rsAUX.Close
        Else
            ArrReporte(i, MaxRow) = 0
        End If
        
        i = i + 1
        If Trim(rs!tipotrabajador) = "01" Then
            ArrReporte(i, MaxRow) = Round((totaportes / 6) * CantMes, 2)
        Else
            ArrReporte(i, MaxRow) = Round(((totaportes * 30) / 6) * CantMes, 2)
        End If
        i = i + 1
        ArrReporte(i, MaxRow) = Abs(ArrReporte(i - 1, MaxRow) - ArrReporte(i - 2, MaxRow))
        
        MaxRow = MaxRow + 1
        
        'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
        sSQLI = ""
        For i = 1 To 50
            Campo = "I" & Format(i, "00")
            sCol = BuscaColumna(Campo, MaxCol)
            If sCol > 0 Then
                sSQLI = sSQLI & IIf(Len(Trim(ArrReporte(sCol, MaxRow - 1))) = 0, "0", ArrReporte(sCol, MaxRow - 1)) & ","
            Else
                sSQLI = sSQLI & "0,"
            End If
        Next
        
        'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
        sSQLP = ""
        For i = 1 To 50
            Campo = "P" & Format(i, "00")
            sCol = BuscaColumna(Campo, MaxCol)
            If sCol > 0 Then
                sSQLP = sSQLP & ArrReporte(sCol, MaxRow - 1) & ","
            Else
                sSQLP = sSQLP & "0,"
            End If
        Next
        
        sSQL = ""
        sSQL = "INSERT plaprovgrati VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MaxRow - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','',"
        sSQL = sSQL & " '','',''," & sSQLI & sSQLP & ArrReporte(MaxColTemp + 1, MaxRow - 1) & ",0,"
        sSQL = sSQL & ArrReporte(MaxColTemp + 2, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 3, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 4, MaxRow - 1) & ",'" & Format(FecProceso, "DD/MM/YYYY") & "',GETDATE(),'" & wuser & "','" & ArrReporte(COL_AREA, MaxRow - 1) & "',' ')"
        
        cn.Execute (sSQL)
        
        totaportes = 0
        rs.MoveNext
    Loop
    
End If
    Carga_Prov_Grati
End Sub


Private Function CantidadMesesCalculo(ByVal pFecIngreso) As String
Dim mesestmp As String
If Year(pFecIngreso) < Txtano.Text Then
    If Cmbmes.ListIndex + 1 < 7 Then
        mesestmp = Cmbmes.ListIndex + 1
    Else
        mesestmp = (Cmbmes.ListIndex + 1) - 6
    End If
Else
    
    If Cmbmes.ListIndex + 1 >= Month(pFecIngreso) Then
        If Cmbmes.ListIndex + 1 < 7 Then
            mesestmp = Cmbmes.ListIndex + 1
            If Month(pFecIngreso) > 1 Then mesestmp = mesestmp - (Month(pFecIngreso) - 1)
            If Day(pFecIngreso) <> 1 Then mesestmp = mesestmp - 1
        Else
            mesestmp = Cmbmes.ListIndex + 1 - 6
            If Month(pFecIngreso) > 7 Then mesestmp = mesestmp - ((Month(pFecIngreso) - 6) - 1)
            If Month(pFecIngreso) > 6 Then
                If Day(pFecIngreso) <> 1 Then mesestmp = mesestmp - 1
            End If
        End If
    Else
        mesestmp = 0
    End If
    
End If
CantidadMesesCalculo = mesestmp

End Function
