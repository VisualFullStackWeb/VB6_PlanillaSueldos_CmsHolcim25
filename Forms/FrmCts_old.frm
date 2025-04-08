VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depositos de CTS"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "FrmCts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Certificados"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reporte por Bancos"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte CTS"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   7455
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmCts.frx":030A
         Left            =   840
         List            =   "FrmCts.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Txtano 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   100
         Width           =   255
         Size            =   "450;661"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deposito"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   315
         Left            =   6120
         TabIndex        =   13
         Top             =   120
         Width           =   1095
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
         TabIndex        =   6
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5535
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   7455
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "FrmCts.frx":039A
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9340
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
            DataField       =   "total"
            Caption         =   "Rem. Afecta"
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
            DataField       =   "totliquid"
            Caption         =   "Deposito"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "fingreso"
            Caption         =   "fingreso"
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
               ColumnWidth     =   4649.953
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
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
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
   Begin MSAdodcLib.Adodc AdoCabeza 
      Height          =   330
      Left            =   2520
      Top             =   4080
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
Attribute VB_Name = "FrmCts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mperiodo As String
Dim SQL As String
Dim mpag As Integer
Dim ruta As String
Dim mlinea As Integer

Private Sub Cmbmes_Click()
Carga_Cts
End Sub

Private Sub Command1_Click()
Reporte_Cts
End Sub

Private Sub Command2_Click()
Reporte_Cts_Bancos
End Sub

Private Sub Command3_Click()
Procesa_Certifica_Cts
End Sub

Private Sub Command4_Click()
Calculo_CTS
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7665
Me.Width = 7575
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
If Month(Date) = 1 Then
   Txtano.Text = Format(Year(Date) - 1, "0000")
   Cmbmes.ListIndex = 11
Else
  Txtano.Text = Format(Year(Date), "0000")
  Cmbmes.ListIndex = Month(Date) - 2
End If
   
End Sub

Private Sub Calculo_CTS()
Dim mcad As String
Dim mcadIns As String
Dim mCadVal As String
Dim mCadFields As String
Dim mFactor As Currency
Dim mTotal As Currency
Dim mDepo As Currency
Dim rs2 As ADODB.Recordset
Dim rsAfectos As ADODB.Recordset
Dim i As Integer, X As Integer
Dim INSCAD As String, VALUESCAD As String
Dim NRO As Integer, cad As String
Dim VALORES(50) As Integer

Screen.MousePointer = vbArrowHourglass

mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
SQL = "select factorcts from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, SQL)) Then
   If IsNull(rs!factorcts) Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If rs!factorcts = 0 Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
Else
   MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
   Screen.MousePointer = vbDefault
   Exit Sub
End If

mFactor = rs!factorcts
rs.Close

SQL = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='C' and status<>'*' order by cod_remu"
If (fAbrRst(rsAfectos, SQL)) Then
   rsAfectos.MoveFirst
Else
  MsgBox "No Se Registran Remuneraciones Afectas para el Calculo", vbCritical, "Calculo de CTS"
  Screen.MousePointer = vbDefault
  Exit Sub
End If

mcad = ""
mCadFields = ""

NRO = 0
Do While Not rsAfectos.EOF
   NRO = NRO + 1
   mcad = mcad & "SUM(I" & Format(rsAfectos!cod_remu, "00") & "),"
   mCadFields = mCadFields & "I" & Format(rsAfectos!cod_remu, "00") & ","
   rsAfectos.MoveNext
Loop

mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
SQL = "select " & mcad
SQL = SQL & ",placod from plahistorico where cia='" & _
wcia & "' and year(fechaproceso)=" & Trim(Txtano.Text) & _
" and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & _
" and status<>'*' Group by placod"

  If (fAbrRst(rs, SQL)) Then
     rs.MoveFirst
  Else
     MsgBox "No Se Registran Boletas para el Calculo", vbCritical, "Calculo de CTS"
     Screen.MousePointer = vbDefault
     Exit Sub
  End If

SQL = wInicioTrans
cn.Execute SQL

SQL = "update platserdep set status='*' where cia='" & wcia & "' and periodo='" & mperiodo & "' and status<>'*'"
cn.Execute SQL

Do While Not rs.EOF
   SQL = "select fingreso,cargo,ctsbanco,ctstipcta,ctsmoneda,ctsnumcta from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
   If (fAbrRst(rs2, SQL)) Then rs2.MoveFirst
   rsAfectos.MoveFirst
   mcadIns = ""
   i = 0
   mTotal = 0: mDepo = 0
   Do While Not rsAfectos.EOF
      mCadVal = "0"
      mCadVal = rs(i)
      
      If Not IsNull(rs(i)) Then mTotal = mTotal + rs(i)
      VALORES(Val(rsAfectos(0))) = rs(i)
      mcadIns = mcadIns & mCadVal & ","
      rsAfectos.MoveNext
      i = i + 1
   Loop
   mDepo = Round((mTotal * mFactor) / 100, 2)
  
   INSCAD = "i05,i07,i08," & _
   "i16,i17,i18,i19,i20," & _
   "i21,i22,i23,i24,i25,i26,i27,i28,i29,i30,i31," & _
   "i32,i33,i34,i35,i36,i37,i38,i39,i40,i41,i42," & _
   "i43,i44,i45,i46,i47,i48,i49,i50"
   
   INSCAD = ""
   VALUESCAD = ""
   For X = 1 To 50
       INSCAD = INSCAD & "I" & Format(X, "00") & ","
       VALUESCAD = VALUESCAD & VALORES(X) & ","
   Next
   
   'VALUESCAD = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
  
   mcadIns = mcadIns & Str(mTotal) & ","
   SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
   SQL = SQL & "insert into platserdep(cia,placod,fechaingreso,cargo," & INSCAD
   SQL = SQL & "total,factor,fecha,periodo,banco,cta,moneda,nro_cta,status,interes,TOTLIQUID) "
   SQL = SQL & "values('" & wcia & "','" & Trim(rs!PLACOD) & "','" & Format(rs2!fingreso, FormatFecha) & "','" & Trim(rs2!cargo) & "',"
   SQL = SQL & VALUESCAD & mTotal & ","
   SQL = SQL & "" & mFactor & "," & FechaSys & ",'" & mperiodo & "','" & rs2!ctsbanco & "','" & rs2!ctstipcta & "',"
   SQL = SQL & "'" & rs2!ctsmoneda & "','" & _
   Trim(rs2!ctsnumcta) & "','',0,0" & ")"
   
   
'   SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
'   SQL = SQL & "insert into platserdep(cia,placod,fechaingreso,cargo," & mCadFields
'   SQL = SQL & "total,factor,fecha,periodo,banco,cta,moneda,nro_cta,status,interes,TOTLIQUID," & INSCAD & ") "
'   SQL = SQL & "values('" & wcia & "','" & Trim(RS!PLACOD) & "','" & Format(rs2!fingreso, FormatFecha) & "','" & Trim(rs2!cargo) & "',"
'   SQL = SQL & mcadIns
'   SQL = SQL & "" & mFactor & "," & FechaSys & ",'" & mperiodo & "','" & rs2!ctsbanco & "','" & rs2!ctstipcta & "',"
'   SQL = SQL & "'" & rs2!ctsmoneda & "','" & _
'   Trim(rs2!ctsnumcta) & "','',0,0" & "," & VALUESCAD & ")"
   
   cn.Execute SQL
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
Loop

SQL = wFinTrans
cn.Execute SQL

If rs.State = 1 Then rs.Close
If rsAfectos.State = 1 Then rsAfectos.Close
Screen.MousePointer = vbDefault
Carga_Cts
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
Carga_Cts
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Carga_Cts()
If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
mperiodo = Trim(Txtano.Text) & Format(Cmbmes.ListIndex + 1, "00")
SQL = nombre()
SQL = SQL & "a.placod,a.total,a.totliquid,b.fingreso " _
& "from platserdep a,planillas b " _
& "where a.cia='" & wcia & "' and a.periodo='" & mperiodo & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by nombre"

cn.CursorLocation = adUseClient
Set AdoCabeza.Recordset = cn.Execute(SQL$, 64)
If AdoCabeza.Recordset.RecordCount > 0 Then
   AdoCabeza.Recordset.MoveFirst
   Command4.Enabled = False
   SQL = "select SUM(totliquid) from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and status<>'*'"
   If (fAbrRst(rs, SQL)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   Lbltotal.Caption = "0.00"
End If
Dgrdcabeza.Refresh
End Sub
Public Sub Elimina_Cts()

mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")

If MsgBox("Desea Eliminar Calculo de Cts ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    SQL = wInicioTrans
    cn.Execute SQL
    
    SQL = "update platserdep set status='*' where cia='" & wcia & "' and periodo='" & mperiodo & "' and status<>'*'"
    cn.Execute SQL
    
    SQL = wFinTrans
    cn.Execute SQL
    
    Carga_Cts
End If

Screen.MousePointer = vbDefault

End Sub
Private Sub Reporte_Cts()
Dim mcad As String
Dim mtotd As Currency
Dim mtott As Currency
mpag = 0
If AdoCabeza.Recordset.RecordCount <= 0 Then
   MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
   Exit Sub
End If
ruta = App.Path & "\REPORTS\" & "RepCts.txt"
Open ruta For Output As #1
Cabeza_Lista_CTS
mtotd = 0: mtott = 0
AdoCabeza.Recordset.MoveFirst
Do While Not AdoCabeza.Recordset.EOF
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Cabeza_Lista_CTS
   mcad = lentexto(55, Left(AdoCabeza.Recordset!nombre, 55))
   mcad = AdoCabeza.Recordset!PLACOD & "   " & mcad & Space(5) & fCadNum(AdoCabeza.Recordset!Total, "##,###,##0.00") & Space(5) & fCadNum(AdoCabeza.Recordset!totliquid, "##,###,##0.00")
   Print #1, Space(2) & mcad
   mtotd = mtotd + AdoCabeza.Recordset!totliquid
   mtott = mtott + AdoCabeza.Recordset!Total
   AdoCabeza.Recordset.MoveNext
Loop
Print #1,
Print #1, Space(35) & "TOTAL :                           " & fCadNum(mtott, "###,###,##0.00") & Space(4) & fCadNum(mtotd, "###,###,##0.00")
Close #1
Call Imprime_Txt("RepCts.txt", ruta)
End Sub
Private Sub Cabeza_Lista_CTS()
mpag = mpag + 1
Print #1, Chr(18) & Space(2) & Trim(Cmbcia.Text) & Space(25) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(22) & "DEPOSITO POR TIEMPO DE SERVICIO"
Print #1, Space(23) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000") & Chr(15)
Print #1, Space(2) & String(107, "-")
Print #1, Space(2) & "CODIGO            NOMBRE                                               REMUNERACION          MONTO A "
Print #1, Space(2) & "                                                                        COMPUTABLE          DEPOSITAR"
Print #1, Space(2) & String(107, "-")
mlinea = 10
End Sub
Private Sub Reporte_Cts_Bancos()
Dim mcad As String
Dim mcod As String
Dim mItem As Integer
Dim mTotDep As Currency
Dim rs2 As ADODB.Recordset
Dim wciamae As String
Dim mtotsbco As Currency
Dim mtotnbco As Integer
Dim mtotsbcoTS As Currency
Dim mtotnbcoTS As Integer
Dim mtotsbcoTD As Currency
Dim mtotnbcoTD As Integer

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")

Select Case StrConv(gsAdminDB, 1)
       Case Is = "MYSQL"
            SQL$ = "CONCAT(rtrim(nom_1),' ',rtrim(nom_2)) AS nombre,"
       Case Is = "SQL SERVER"
            SQL$ = "rtrim(nom_1)+' '+rtrim(nom_2) as nombre,"
End Select

SQL = "select ap_pat,ap_mat," & SQL
SQL = SQL & "a.placod,a.total,a.totliquid,a.banco,a.nro_cta,a.moneda,b.fingreso " _
& "from platserdep a,planillas b " _
& "where a.cia='" & wcia & "' and a.periodo='" & mperiodo & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by banco,B.moneda,nombre"


If (fAbrRst(rs, SQL)) Then
   rs.MoveFirst
Else
   MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
   Exit Sub
End If

ruta = App.Path & "\REPORTS\" & "BcoCts.txt"
Open ruta For Output As #1
Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
mcod = rs!banco & rs!moneda
mItem = 0: mTotDep = 0
Do While Not rs.EOF
   If mcod <> rs!banco & rs!moneda Then
      Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
      Print #1, Space(37) & "TOTAL INSCRITOS < " & fCadNum(mItem, "#####") & " >                            TOTAL ACUM.       < " & fCadNum(mTotDep, "##,###,##0.00") & " >"
      mItem = 0: mTotDep = 0
     Print #1, Chr(12) + Chr(13): Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
     mcod = rs!banco & rs!moneda
   End If
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
   mcad = lentexto(20, Left(rs!nro_cta, 20)) & Space(5)
   mcad = mcad & lentexto(20, Left(rs!ap_pat, 20)) & Space(2) & lentexto(20, Left(rs!ap_mat, 20)) & Space(2) & Space(2) & lentexto(25, Left(rs!nombre, 25))
   mcad = mcad & Space(2) & fCadNum(rs!totliquid, "##,###,##0.00")
   Print #1, Space(12) & mcad
   mlinea = mlinea + 1
   mItem = mItem + 1
   mTotDep = mTotDep + rs!totliquid
   rs.MoveNext
Loop
Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
Print #1, Space(37) & "TOTAL INSCRITOS < " & fCadNum(mItem, "#####") & " >                            TOTAL ACUM.       < " & fCadNum(mTotDep, "##,###,##0.00") & " >"
If rs.State = 1 Then rs.Close

SQL = "select distinct(banco) from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  status<>'*'"
If (fAbrRst(rs, SQL)) Then rs.MoveFirst
Print #1, Chr(12) + Chr(13): Cabeza_Banco_Resumen
mtotnbcoTS = 0: mtotsbcoTS = 0
mtotnbcoTD = 0: mtotsbcoTD = 0
Do While Not rs.EOF
   mcad = ""
   wciamae = Determina_Maestro("01007")
   mtotnbco = 0: mtotsbco = 0
   SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs!banco & "'" & Space(5)
   SQL$ = SQL$ & wciamae
   If (fAbrRst(rs2, SQL$)) Then mcad = lentexto(33, Left(rs2!descrip, 23)) Else mcad = Space(33)
   If rs2.State = 1 Then rs2.Close
   
   'Dolares
   SQL = "select count(placod) as num from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  banco='" & rs!banco & "' and moneda<>'" & wmoncont & "' and status<>'*'"
   If (fAbrRst(rs2, SQL)) Then
      If rs2!num = 0 Then
         mcad = mcad & "     0"
      Else
         mcad = mcad & " " & fCadNum(rs2!num, "#####")
      End If
      mtotnbco = mtotnbco + rs2!num
      mtotnbcoTD = mtotnbcoTD + rs2!num
      rs2.Close
      SQL = "select sum(totliquid) as depo from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  banco='" & rs!banco & "' and moneda<>'" & wmoncont & "' and status<>'*'"
      If (fAbrRst(rs2, SQL)) Then
         If IsNull(rs2!depo) Then
            mcad = mcad & " " & fCadNum(0, "#,###,##0.00")
         Else
            mcad = mcad & " " & fCadNum(rs2!depo, "#,###,##0.00")
            mtotsbco = mtotsbco + rs2!depo
            mtotsbcoTD = mtotsbcoTD + rs2!depo
         End If
      End If
      rs2.Close
      mcad = mcad & Space(2)
   End If
   'Soles
   SQL = "select count(placod) as num from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  banco='" & rs!banco & "' and moneda='" & wmoncont & "' and status<>'*'"
   If (fAbrRst(rs2, SQL)) Then
      mcad = mcad & Space(8)
      If rs2!num = 0 Then
         mcad = mcad & "     0"
      Else
         mcad = mcad & " " & fCadNum(rs2!num, "#####")
      End If
      mtotnbco = mtotnbco + rs2!num
      mtotnbcoTS = mtotnbcoTS + rs2!num
      rs2.Close
      SQL = "select sum(totliquid) as depo from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  banco='" & rs!banco & "' and moneda='" & wmoncont & "' and status<>'*'"
      If (fAbrRst(rs2, SQL)) Then
         If IsNull(rs2!depo) Then
            mcad = mcad & " " & fCadNum(0, "#,###,##0.00")
         Else
            mcad = mcad & " " & fCadNum(rs2!depo, "#,###,##0.00")
            mtotsbco = mtotsbco + rs2!depo
            mtotsbcoTS = mtotsbcoTS + rs2!depo
         End If
      End If
      rs2.Close
      'Total
      mcad = mcad & Space(10) & " " & fCadNum(mtotnbco, "#####") & " " & fCadNum(mtotsbco, "#,###,##0.00")
      mcad = mcad & Space(2)
   End If
   Print #1, Space(12) & mcad
   rs.MoveNext
Loop
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
'Total General
mcad = "  ** T O T A L ** " & Space(16) & fCadNum(mtotnbcoTD, "#####") & " " & fCadNum(mtotsbcoTD, "#,###,##0.00")
mcad = mcad & Space(11) & fCadNum(mtotnbcoTS, "#####") & " " & fCadNum(mtotsbcoTS, "#,###,##0.00")
mcad = mcad & Space(11) & fCadNum(mtotnbcoTS + mtotnbcoTD, "#####") & " " & fCadNum(mtotsbcoTS + mtotsbcoTD, "#,###,##0.00")
Print #1, Space(12) & mcad
Close #1
Call Imprime_Txt("BcoCts.txt", ruta)
End Sub
Private Sub Cabeza_Banco_CTS(banco, moneda)
Dim wciamae As String
Dim rs2 As ADODB.Recordset
Dim mBanc As String
Dim MMON As String
If Trim(banco) = "" Then
   mBanc = "SIN BANCO"
Else
   mBanc = ""
   wciamae = Determina_Maestro("01007")
   SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & banco & "'" & Space(5)
   SQL$ = SQL$ & wciamae
   If (fAbrRst(rs2, SQL$)) Then mBanc = lentexto(30, Left(rs2!descrip, 30))
   If rs2.State = 1 Then rs2.Close
End If
If Trim(moneda) = "" Then
   MMON = ""
ElseIf moneda = wmoncont Then
   MMON = "Moneda Nacional (" & moneda & ")"
Else
   MMON = "Moneda Extranjera (" & moneda & ")"
End If
Print #1, Chr(18) & Space(10) & Trim(Cmbcia.Text)
Print #1,
Print #1, Space(27) & mBanc
Print #1,
Print #1, Space(27) & MMON & Chr(15)
Print #1,
Print #1, Space(16) & "CUENTA                              APELLIDOS Y NOMBRES                                              MONTO"
Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
Print #1,
mlinea = 9
End Sub
Private Sub Cabeza_Banco_Resumen()
Print #1, Chr(18) & Space(10) & Trim(Cmbcia.Text)
Print #1,
Print #1, Space(27) & "R E S U M E N" & Chr(15)
Print #1,
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
Print #1,
Print #1, Space(46) & "DEP.MON.EXT (DOLAR)              DEP. SOLES                   T O T A L"
Print #1, Space(16) & "BANCOS                        -------------------          -------------------          -------------------"
Print #1, Space(46) & "# TRAB    MONTO S/.          # TRAB    MONTO S/.          # TRAB    MONTO S/."
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
Print #1,
mlinea = 9
End Sub
Private Sub Procesa_Certifica_Cts()
Dim mItem As Integer
Dim mcad As String
Dim mnombre As String
Dim cadnombre As String
Dim mcadrem As String
Dim mtot As Currency
Dim mFactor As Currency
Dim mdir As String
Dim mmoneda As String
Dim cadefectivo As String
Dim mdpto As String
Dim rs2 As ADODB.Recordset
Dim rsRem As ADODB.Recordset
Dim wciamae As String
Dim mempleado As String
Dim mobrero As String
Dim mfec As String
Dim i As Integer

SQL = "select factorcts from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, SQL)) Then
   If IsNull(rs!factorcts) Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If rs!factorcts = 0 Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
Else
   MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
   Screen.MousePointer = vbDefault
   Exit Sub
End If
mFactor = rs!factorcts
rs.Close

mfec = fMaxDay(Cmbmes.ListIndex + 1, Val(Txtano)) & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano
cadnombre = nombre()
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
mdir = ""
mdpto = ""
SQL$ = "select distinct(direcc),nro,u.dpto  from cia c,ubigeos u where c.cod_cia='01' and c.status<>'*' and left(u.cod_ubi,5)=left(c.cod_ubi,5)"
If (fAbrRst(rs, SQL$)) Then mdir = Trim(rs!direcc) & " No " & Trim(rs!NRO): mdpto = Trim(rs!DPTO)
rs.Close

SQL$ = "Select * from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and status<>'*' order by placod"
If Not (fAbrRst(rs, SQL$)) Then Else rs.MoveFirst

ruta = App.Path & "\REPORTS\" & "CertCts.txt"
Open ruta For Output As #1

Do While Not rs.EOF
   wciamae = Determina_Maestro("01007")
   SQL$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs!banco & "'" & Space(5)
   SQL$ = SQL$ & wciamae
   If (fAbrRst(rs2, SQL$)) Then mcad = lentexto(48, Left(rs2!descrip, 48)) Else mcad = Space(48)
   If rs2.State = 1 Then rs2.Close

   If rs!moneda = wmoncont Then
      mmoneda = "     CUENTA A PLAZO/MONEDA NACIONAL (SOLES)       "
   Else
      mmoneda = "     CUENTA A PLAZO/EXTRANJERA NACIONAL (DOLARES) "
   End If
   SQL$ = cadnombre & "placod,tipotrabajador,fingreso,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, SQL$)) Then mnombre = lentexto(41, Left(rs2!nombre, 41)) Else mnombre = Space(40)
   If rs2!tipotrabajador = "01" Then
      mempleado = " EMPLEADO [X] "
      mobrero = " OBRERO   [ ] "
   Else
      mempleado = " EMPLEADO [ ] "
      mobrero = " OBRERO   [X] "
   End If
   
    Print #1, Chr(18) & "Liquidacion de Compensacion por Tiempo de Servicios(CTS)" & Chr(15)
    Print #1, "DECRETO DE URGENCIA No 127 - 2000 (30.12.2000) DECRETO SUPREMO No 001-2001-TR (22.01.2001)"
    Print #1, Chr(218) & String(87, Chr(196)) & Chr(194) & String(44, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "Nombre o Razon Social del Empleador :  " & lentexto(48, Left(Cmbcia.Text, 48)) & Chr(179) & "Ciudad y Fecha  " & lentexto(16, Left(mdpto, 16)) & Format(Date, "dd/mm/yyyy") & "  " & Chr(179)
    Print #1, Chr(195) & String(77, Chr(196)) & Chr(194) & String(9, Chr(196)) & Chr(193) & String(16, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "Direccion del Empleador :   " & lentexto(49, Left(mdir, 49)) & Chr(179) & "  Deposito  [X]           " & Chr(179) & "     Pago Directo  [ ]     " & Chr(179)
    Print #1, Chr(195) & String(53, Chr(196)) & Chr(194) & String(23, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(15) & "ENTIDAD DEPOSITARIA" & Space(19) & Chr(179) & "Tipo de Cuenta  " & Space(34) & Chr(179) & "No de Cuenta" & Space(15) & Chr(179)
    Print #1, Chr(179) & Space(5) & mcad & Chr(179) & mmoneda & Chr(179) & Space(1) & lentexto(26, Left(rs!nro_cta, 26)) & Chr(179)
    Print #1, Chr(195) & String(53, Chr(196)) & Chr(193) & String(23, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(28) & "DATOS DEL TRABAJADOR" & Space(29) & Chr(179) & "    CONCEPTOS REMUNERATIVOS PARA DETERMINAR LA        " & Chr(179)
    Print #1, Chr(179) & Space(77) & Chr(179) & "               REMUNERACION COMPUTABLE                " & Chr(179)
    Print #1, Chr(195) & String(17, Chr(196)) & Chr(194) & String(59, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "CODIGO: " & lentexto(9, Left(rs!PLACOD, 9)) & Chr(179) & "Apell.y Nombres : " & mnombre & Chr(179) & "         CONCEPTO         " & Chr(179) & "            MONTO          " & Chr(179)
    Print #1, Chr(195) & String(17, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(194) & String(17, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "Fecha de Ingreso " & Chr(179) & mempleado & Chr(179) & "  Fecha de Cese  " & Chr(179) & "      Motivo de Cese      " & Chr(179) & "                          " & Chr(179) & "                           " & Chr(179)
    mcad = ""
    mItem = 0
    For i = 4 To 53
        If rs(i) <> 0 Then
           mcadrem = Chr(179)
           SQL = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' and codinterno='" & Right(rs(i).Name, 2) & "'"
           If (fAbrRst(rsRem, SQL$)) Then mcadrem = mcadrem & lentexto(26, Left(rsRem!descripcion, 26)) Else mcadrem = mcadrem & Space(26)
           mcadrem = mcadrem & Chr(179) & Space(12) & fCadNum(rs(i), "##,###,##0.00") & Space(2) & Chr(179)
           rsRem.Close
           mItem = mItem + 1
           Select Case mItem
                  Case Is = 1
                       mcad = Chr(179) & "   " & Format(rs2!fingreso, "dd/mm/yyyy") & "    " & Chr(179) & mobrero & Chr(179) & Space(3)
                       If IsNull(rs2!fcese) Then mcad = mcad & Space(10) Else mcad = mcad & Format(rs2!fcese, "dd/mm/yyyy")
                       mcad = mcad & Space(4) & Chr(179) & Space(26) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                       Print #1, Chr(179) & "   " & Chr(179) & Space(19) & "PERIODO DE SERVICIOS QUE SE CANCELA" & Space(19) & Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
                  Case Is = 2
                       If Month(rs2!fingreso) = Cmbmes.ListIndex + 1 And Year(rs2!fingreso) = Val(Txtano.Text) Then
                          mcad = "Del : " & Format(rs2!fingreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(4) & "No De Meses :     No de Dias :  " & fCadNum((Val(Left(mfec, 2))) - (Day(rs2!fingreso)), "##") & Space(2) & mcadrem
                          cadefectivo = Chr(179) & "    POR" & Space(5) & fCadNum((Val(Left(mfec, 2))) - (Day(rs2!fingreso)), "##") & Space(5) & "Dias" & Space(9) & Chr(179)
                       Else
                          mcad = "Del : " & "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "  Al : " & mfec & Space(4) & "No De Meses :   1   No de Dias :  " & Space(2) & mcadrem
                          cadefectivo = Chr(179) & "    POR" & Space(5) & " 1" & Space(5) & "Mes " & Space(9) & Chr(179)
                       End If
                       mcad = Chr(179) & " 1 " & Chr(179) & mcad
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(197) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 3
                       mcad = Chr(179) & " 2 " & Chr(179) & "Periodo no Comp.(desc.en Obser.)" & Space(41) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(193) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 4
                       mcad = Chr(179) & "TIEMPO EFECTIVO A LIQUIDAR (1-2)" & Space(45) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 5
                       mcad = Chr(179) & "REM.COMPUT.MENSUAL PARA LIQUID.EL PER." & Space(39) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 6
                       mcad = Chr(179) & "Un dozavo de la Rem.compt.mensual     " & Space(39) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 7
                       mcad = Chr(179) & "Un 30avo del 12avo de la Rem. Compt.  " & Space(39) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 8
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 9
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 10
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
           End Select
        End If
    Next
    For i = mItem To 10
        mItem = mItem + 1
        mcadrem = Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
        Select Case mItem
               Case Is = 1
                    mcad = Chr(179) & "   " & Format(rs2!fingreso, "dd/mm/yyyy") & "    " & Chr(179) & mobrero & Chr(179) & Space(3)
                    If IsNull(rs2!fcese) Then mcad = mcad & Space(10) Else mcad = mcad & Format(rs2!fcese, "dd/mm/yyyy")
                    mcad = mcad & Space(4) & Chr(179) & Space(26) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                    Print #1, Chr(179) & "   " & Chr(179) & Space(19) & "PERIODO DE SERVICIOS QUE SE CANCELA" & Space(19) & Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
               Case Is = 2
                    If Month(rs2!fingreso) = Cmbmes.ListIndex + 1 And Year(rs2!fingreso) = Val(Txtano.Text) Then
                       mcad = "Del : " & Format(rs2!fingreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(4) & "No De Meses :     No de Dias :  " & fCadNum((Val(Left(mfec, 2))) - (Day(rs2!fingreso)), "##") & Space(2) & mcadrem
                       cadefectivo = Chr(179) & "    POR" & Space(5) & fCadNum((Val(Left(mfec, 2))) - (Day(rs2!fingreso)), "##") & Space(5) & "Dias" & Space(9) & Chr(179)
                    Else
                       mcad = "Del : " & "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "  Al : " & mfec & Space(4) & "No De Meses :   1   No de Dias :  " & Space(2) & mcadrem
                       cadefectivo = Chr(179) & "    POR" & Space(5) & " 1" & Space(5) & "Mes " & Space(9) & Chr(179)
                    End If
                    mcad = Chr(179) & " 1 " & Chr(179) & mcad
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(197) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 3
                    mcad = Chr(179) & " 2 " & Chr(179) & "Periodo no Comp.(desc.en Obser.)" & Space(41) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(193) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 4
                    mcad = Chr(179) & "TIEMPO EFECTIVO A LIQUIDAR (1-2)" & Space(45) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 5
                    mcad = Chr(179) & "REM.COMPUT.MENSUAL PARA LIQUID.EL PER." & Space(39) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 6
                    mcad = Chr(179) & "Un dozavo de la Rem.compt.mensual     " & Space(39) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 7
                    mcad = Chr(179) & "Un 30avo del 12avo de la Rem. Compt.  " & Space(39) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 8
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 9
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 10
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
        End Select
    Next
    mcadrem = Chr(179)
    SQL = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' and codinterno='" & Right(rs(i).Name, 2) & "'"
    mcadrem = mcadrem & "TOTAL" & Space(21)
    mcadrem = mcadrem & Chr(179) & Space(12) & fCadNum(rs!Total, "##,###,##0.00") & Space(2) & Chr(179)
    mcad = Chr(179) & Space(77) & mcadrem
    Print #1, mcad
    Print #1, Chr(195) & String(77, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(43) & "LIQUIDACION DE LAS CTS CON EFECTO CANCELATORIO" & Space(43) & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(194) & String(71, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(3) & "TIEMPO EFECTIVO A LIQUIDAR" & Space(3) & Chr(179) & Space(27) & "CALCULO DE LA CTS" & Space(27) & Chr(179) & "            MONTO          " & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    mcad = cadefectivo & Space(24) & fCadNum(rs!Total, "##,###,##0.00") & "  x " & fCadNum(mFactor, "##0.00") & Space(24) & Chr(179) & Space(27) & Chr(179)
    Print #1, mcad
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(32) & Chr(179) & Space(20) & "TOTAL CTS DEPOSITADA O PAGADA :" & Space(20) & Chr(179) & Space(12) & fCadNum(rs!totliquid, "##,###,##0.00") & Space(2) & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(194) & String(43, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "OBSERVACIONES :" & Space(45) & Chr(179) & " CONSTANCIA DE RECEPCION :" & Space(45) & Chr(179)
    Print #1, Chr(179) & Space(60) & Chr(179) & " 1)De la Presente Liquidacion" & Space(42) & Chr(179)
    Print #1, Chr(179) & Space(60) & Chr(179) & " 2)Del Documento que Acredita el Deposito de la CTS antes detalleda" & Space(4) & Chr(179)
    Print #1, Chr(195) & String(11, Chr(196)) & Chr(194) & String(48, Chr(196)) & Chr(180) & Space(71) & Chr(179)
    Print #1, Chr(179) & "VoB." & Space(7) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & String(48, "-") & Chr(179) & String(71, "-") & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & "Nombres,Apellidos y Cargo del Representante del " & Chr(179) & Space(17) & "Firma del Trabajador (Huella Digital)" & Space(17) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & "            Empleador(Sello y Firma)            " & Chr(179) & Space(23) & mnombre & Space(7) & Chr(179)
    Print #1, Chr(192) & String(11, Chr(196)) & Chr(193) & String(48, Chr(196)) & Chr(193) & String(71, Chr(196)) & Chr(217)
    
    Print #1, SaltaPag
   rs.MoveNext
Loop
Close #1
Call Imprime_Txt("CertCts.txt", ruta$)

End Sub

Private Sub PROVICIONES_CTS()
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

Const COL_PRIMERMES = 5
Const COL_PRIMERMES = 11

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

sSQL = "SELECT concepto,campo,sn_promedio,0 AS factor,sn_carga FROM estructura_provisiones WHERE tipo='C' and "
sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
sSQL = sSQL & " UNION ALL SELECT concepto,campo,sn_promedio,b.factor,sn_carga FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.codinterno=a.campo) WHERE a.tipo='C'"

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        MaxCol = MaxCol + 1
        rs.MoveNext
    Loop
    rs.MoveFirst
    MaxColTemp = MaxCol + 1
    MaxCol = MaxCol + 8
    
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
        
        ' PROMEDIO GRATIFICACION
        i = MaxColTemp
        ArrReporte(i, 0) = "P15"
        If Cmbmes.ListIndex + 1 > 6 Then
            SQL = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=7 and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
        Else
            SQL = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=12 and month(fechaproceso)=" & Txtano.Text - 1 & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
        End If
        
        If (fAbrRst(rsAUX, SQL)) Then
            ArrReporte(i, MaxRow) = rsAUX(0) / 180
            rsAUX.Close
        Else
            ArrReporte(i, MaxRow) = 0
        End If
        
        
        'PROVISION DEL AÑO PASADO
        i = i + 1
        If Cmbmes.ListIndex <= COL_PRIMERMES Then
            SQL = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=7 and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
            If (fAbrRst(rsAUX, SQL)) Then
                ArrReporte(i, MaxRow) = rsAUX(0)
                rsAUX.Close
            Else
                ArrReporte(i, MaxRow) = 0
            End If
        Else
            SQL = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=7 and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
        End If
                
        
        
        i = i + 1
        
        
        
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
        sSQL = "INSERT plaprovcts VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MaxRow - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','',"
        sSQL = sSQL & " '','',''," & sSQLI & sSQLP & ArrReporte(MaxColTemp + 1, MaxRow - 1) & ",0,"
        sSQL = sSQL & ArrReporte(MaxColTemp + 2, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 3, MaxRow - 1) & "," & ArrReporte(MaxColTemp + 4, MaxRow - 1) & ",'" & Format(FecProceso, "DD/MM/YYYY") & "',GETDATE(),'" & wuser & "','" & ArrReporte(COL_AREA, MaxRow - 1) & "',' ')"
        
        cn.Execute (sSQL)
        
        totaportes = 0
        rs.MoveNext
    Loop
    
End If
    Carga_Prov_Grati
End Sub

