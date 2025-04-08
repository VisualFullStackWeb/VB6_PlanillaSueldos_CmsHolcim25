VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmcarga 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de datos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Plamas"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Historico"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Urbanizaciones (tablas)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Turnos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cuentas Bancarias"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ubigeos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Constantes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carga maestros"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Data Datplamas 
      Caption         =   "Data1"
      Connect         =   "dBASE IV;"
      DatabaseName    =   "C:\Temp\CARGAPLA"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PLAMAST"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Dathistoric 
      Caption         =   "Data1"
      Connect         =   "dBASE IV;"
      DatabaseName    =   "C:\Temp\CARGAPLA"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "HISTORIC"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "Frmcarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cncarga As New ADODB.Connection
Private Sub Command1_Click()
'MAESTROS

Sql$ = "select * from maestros"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst

Sql$ = wInicioTrans
cn.Execute Sql$

Do While Not rs.EOF
   Sql$ = "INSERT INTO maestros values('" & rs!cod_cia & "','" & rs!id_maestro & "','" & rs!ciamaestro & "','" & rs!DESCRIP & "', " _
        & "'" & rs!status & "','" & rs!user_crea & "','" & rs!user_modi & "',now(),now(),'" & rs!Sistema & "','S')"
   cn.Execute Sql$

   rs.MoveNext
Loop

'MAESTROS_2
If rs.State = 1 Then rs.Close
Sql$ = "select * from maestros_2"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst
Do While Not rs.EOF
   Sql$ = "INSERT INTO maestros_2 values('" & rs!ciamaestro & "','" & Trim(rs!COD_MAESTRO2) & "','" & rs!DESCRIP & "', " _
        & "'" & rs!flag1 & "','" & rs!flag2 & "','" & rs!flag3 & "','" & rs!FLAG4 & "','" & rs!flag5 & "','" & rs!status & "','" & rs!user_crea & "', " _
        & "'" & rs!user_modi & "',now(),now(),'" & rs!flag6 & "','" & rs!flag7 & "','" & rs!CODSUNAT & "', " _
        & "'" & rs!generico & "','" & rs!GRUPO & "','" & rs!generico2 & "','" & rs!grupo2 & "')"
        cn.Execute Sql$

   rs.MoveNext
Loop

'MAESTROS_3
If rs.State = 1 Then rs.Close
Sql$ = "select * from maestros_3"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst
Do While Not rs.EOF
   Sql$ = "INSERT INTO maestros_3 values('" & rs!ciamaestro & "','" & Trim(rs!COD_MAESTRO2) & "','" & rs!cod_maestro3 & "','" & rs!DESCRIP & "', " _
        & "'" & rs!flag1 & "','" & rs!flag2 & "','" & rs!status & "','" & rs!user_crea & "', " _
        & "'" & rs!user_modi & "',now(),now(),'" & rs!flag3 & "', " _
        & "'" & rs!generico & "','" & rs!GRUPO & "','" & rs!generico2 & "','" & rs!grupo2 & "')"
        cn.Execute Sql$

   rs.MoveNext
Loop
Sql$ = wFinTrans
cn.Execute Sql$

End Sub
Private Sub Command2_Click()
Sql$ = "select * from placonstante where status<>'*'"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst

Sql$ = wInicioTrans
cn.Execute Sql$
Do While Not rs.EOF
   Sql$ = "INSERT INTO placonstante values('" & rs!cia & "','" & rs!tipomovimiento & "','" & rs!codinterno & "','" & rs!Descripcion & "', " _
        & "'" & rs!deduccion & "','" & rs!aportacion & "','" & rs!personal & "','" & rs!status & "','" & rs!calculo & "','" & rs!adicional & "',now(),'" & rs!Basico & "','01','','','' )"
   cn.Execute Sql$

   rs.MoveNext
Loop
Sql$ = wFinTrans
cn.Execute Sql$
End Sub

Private Sub Command3_Click()
Sql$ = "select * from ubigeos"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst
Sql$ = wInicioTrans
cn.Execute Sql$

Do While Not rs.EOF
   Sql$ = "INSERT INTO ubigeos values('" & rs!cod_ubi & "','" & rs!codpais & "','" & rs!coddpto & "','" & rs!codprov & "','" & rs!CODDIST & "','" & rs!pais & "','" & rs!DPTO & "','" & rs!PROV & "','" & rs!DIST & "','" & rs!cod_postal & "','C') "
   cn.Execute Sql$

   rs.MoveNext
Loop
Sql$ = wFinTrans
cn.Execute Sql$

End Sub

Private Sub Command4_Click()
Sql$ = "select * from bancocta where status<>'*'"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst
Sql$ = wInicioTrans
cn.Execute Sql$
Do While Not rs.EOF
   Sql$ = "INSERT INTO bancocta values('" & rs!cia & "','" & rs!banco & "','" & rs!CUENTA & "','" & rs!moneda & "','" & Format(rs!fecha_crea, FormatFecha) & "','" & rs!user_crea & "',NULL," _
        & "'" & rs!user_modi & "','" & rs!cheque & "','" & rs!status & "','" & rs!bcoasiento & "','" & rs!contable & "','" & rs!ultimo & "','" & rs!aviso & "','" & rs!imprime & "','" & rs!interno & "')"
   cn.Execute Sql$
   rs.MoveNext
Loop
Sql$ = wFinTrans
cn.Execute Sql$
End Sub

Private Sub Command5_Click()
Sql$ = "select * from platurno where status<>'*'"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst
Sql$ = wInicioTrans
cn.Execute Sql$
Do While Not rs.EOF
   Sql$ = "INSERT INTO platurno values('" & rs!cia & "','" & rs!codturno & "','" & rs!Descripcion & "','" & rs!hinicio & "','" & rs!hfinal & "',now(),'" & rs!status & "') "
   cn.Execute Sql$
   rs.MoveNext
Loop
Sql$ = wFinTrans
cn.Execute Sql$
End Sub

Private Sub Command6_Click()
Sql$ = "select * from tablas where status<>'*'"
cncarga.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cncarga.Execute(Sql$, 64)
If rs.RecordCount <= 0 Then
   MsgBox "No Existen elementos a cargar", vbCritical, "Carga de Datos"
   Exit Sub
End If
rs.MoveFirst
Sql$ = wInicioTrans
cn.Execute Sql$
Do While Not rs.EOF
   Sql$ = "INSERT INTO tablas values('" & rs!cod_tipo & "','" & Trim(rs!COD_MAESTRO2) & "','" & rs!DESCRIP & "','" & rs!flag1 & "','" & rs!flag2 & "','" & rs!flag3 & "','" & rs!FLAG4 & "','') "
   cn.Execute Sql$
   rs.MoveNext
Loop
Sql$ = wFinTrans
cn.Execute Sql$
End Sub
Private Sub Command7_Click()
Dim I As Integer
Dim mcad As String
Screen.MousePointer = vbHourglass
Sql$ = wInicioTrans
cn.Execute Sql$
Dathistoric.Recordset.MoveFirst
Num = 1
xItem = 0

Set lRS = Dathistoric.Recordset
Barra.Max = Dathistoric.Recordset.RecordCount
Dim mbol As String
Dim mtrab As String
Do While (Not lRS.EOF)
   Barra.Value = I
   mbol = "": mtrab = ""
   If Mid(lRS!placia, 2, 1) = "P" Then mbol = "04"
   If Mid(lRS!placia, 2, 1) = "N" Then mbol = "01"
   If Mid(lRS!placia, 2, 1) = "V" Then mbol = "02"
   If Mid(lRS!placia, 2, 1) = "G" Then mbol = "03"
   If Mid(lRS!placia, 1, 1) = "O" Then mtrab = "02" Else mtrab = "01"
   
   
   Sql$ = "select p.codauxinterno, p.fcese,p.fingreso,p.codafp,p.numafp,b.importe from planillas p, plaremunbase b " _
        & "where p.placod='" & lRS!PlaCod & "' and b.concepto='01' and p.placod=b.placod"
   If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
   
   Sql = "insert into plahistorico values('01','" & lRS!PlaCod & _
   "','" & rs!codauxinterno & "','" & mbol & "'," & "'" & _
   Format(lRS!pladia, FormatFecha) & "','" & _
   lRS!plasem & "" & "','" & IIf(IsNull(rs!fcese), Null, Format(rs!fcese, FormatFecha)) & _
   "',null,null,'" & IIf(IsNull(lRS!plavacsal), Null, Format(lRS!plavacsal, FormatFecha)) & _
   "','" & IIf(IsNull(lRS!plavacret), Null, Format(lRS!plavacret, FormatFecha)) & _
   "',null," & "'" & Format(rs!fIngreso, FormatFecha) & "','1',"
        
   mcad = "" & IIf(IsNull(lRS!plah01), 0, lRS!plah01) & "," & IIf(IsNull(lRS!plah02), 0, lRS!plah02) & "," & IIf(IsNull(lRS!plah03), 0, lRS!plah03) & "," & IIf(IsNull(lRS!plah04), 0, lRS!plah04) & "," & IIf(IsNull(lRS!plah05), 0, lRS!plah05) & "," & IIf(IsNull(lRS!plah06), 0, lRS!plah06) & "," _
        & "" & IIf(IsNull(lRS!plah20), 0, lRS!plah20) & "," & IIf(IsNull(lRS!plah08), 0, lRS!plah08) & "," & IIf(IsNull(lRS!plah09), 0, lRS!plah09) & "," & IIf(IsNull(lRS!plah10), 0, lRS!plah10) & "," & IIf(IsNull(lRS!plah11), 0, lRS!plah11) & "," & IIf(IsNull(lRS!plah12), 0, lRS!plah12) & "," _
        & "" & IIf(IsNull(lRS!plah13), 0, lRS!plah13) & ",0," & IIf(IsNull(lRS!plah14), 0, lRS!plah14) & "," & IIf(IsNull(lRS!plah16), 0, lRS!plah16) & ",0,0,0,0,0,0," & IIf(IsNull(lRS!plah17), 0, lRS!plah17) & "," & IIf(IsNull(lRS!plah18), 0, lRS!plah18) & "," & IIf(IsNull(lRS!plah19), 0, lRS!plah19) & "," _
        & "" & IIf(IsNull(lRS!plah20), 0, lRS!plah20) & ",0,0,0,0,"
   Sql = Sql & mcad
  
   mcad = "" & IIf(IsNull(lRS!plai01), 0, lRS!plai01) & "," & IIf(IsNull(lRS!plai06), 0, lRS!plai06) & "," & IIf(IsNull(lRS!plai07), 0, lRS!plai07) & "," & IIf(IsNull(lRS!plai08), 0, lRS!plai08) & "," & IIf(IsNull(lRS!plaafp10), 0, lRS!plaafp10) & "," & IIf(IsNull(lRS!plai27), 0, lRS!plai27 - lRS!plaafp10) & "," _
        & "" & IIf(IsNull(lRS!plai04), 0, lRS!plai04) & "," & IIf(IsNull(lRS!plai03), 0, lRS!plai03) & "," & IIf(IsNull(lRS!plai02), 0, lRS!plai02) & "," & IIf(IsNull(lRS!plai10), 0, lRS!plai10) & "," & IIf(IsNull(lRS!plai12), 0, lRS!plai12) & "," & IIf(IsNull(lRS!plai14), 0, lRS!plai14) & "," & IIf(IsNull(lRS!plai15), 0, lRS!plai15) & "," _
        & "" & IIf(IsNull(lRS!plai16), 0, lRS!plai16) & "," & IIf(IsNull(lRS!plai17), 0, lRS!plai17) & "," & IIf(IsNull(lRS!plai18), 0, lRS!plai18) & "," & IIf(IsNull(lRS!plai20), 0, lRS!plai20) & "," & IIf(IsNull(lRS!plai25), 0, lRS!plai25) & ",0,0,0,0,0,0,0,0," & IIf(IsNull(lRS!plai29), 0, lRS!plai29) & "," _
        & "" & IIf(IsNull(lRS!plai30), 0, lRS!plai30) & "," & IIf(IsNull(lRS!plai31), 0, lRS!plai31) & "," & IIf(IsNull(lRS!plai32), 0, lRS!plai32) & ",0,0,0,0,0," & IIf(IsNull(lRS!plai05), 0, lRS!plai05) & "," & IIf(IsNull(lRS!plai13), 0, lRS!plai13) & "," & IIf(IsNull(lRS!plai19), 0, lRS!plai19) & "," & IIf(IsNull(lRS!plai33), 0, lRS!plai33) & "," _
        & "0,0,0,0,0,0,0,0,0,0,0,"
        
   Sql = Sql & mcad
   
   mcad = "" & IIf(IsNull(lRS!plad01), 0, lRS!plad01) & "," & IIf(IsNull(lRS!plad03), 0, lRS!plad03) & ",0," & IIf(IsNull(lRS!plad02), 0, lRS!plad02) & "," & IIf(IsNull(lRS!plad06), 0, lRS!plad06) & "," & IIf(IsNull(lRS!plad12), 0, lRS!plad12) & "," & IIf(IsNull(lRS!plad05), 0, lRS!plad05) & "," & IIf(IsNull(lRS!plad07), 0, lRS!plad07) & "," & IIf(IsNull(lRS!plad04), 0, lRS!plad04) & "," _
        & "" & IIf(IsNull(lRS!plad10), 0, lRS!plad10) & "," & IIf(IsNull(lRS!plad11), 0, lRS!plad11) & ",0," & IIf(IsNull(lRS!plad09), 0, lRS!plad09) & ",0," & IIf(IsNull(lRS!plad08), 0, lRS!plad08) & ",0,0,0,0,0,"
   
   Sql = Sql & mcad
   
   mcad = "" & IIf(IsNull(lRS!plaa01), 0, lRS!plaa01) & "," & IIf(IsNull(lRS!plaa03), 0, lRS!plaa03) & "," & IIf(IsNull(lRS!plaa05), 0, lRS!plaa05) & "," & IIf(IsNull(lRS!plaa02), 0, lRS!plaa02) & ",0,0,0,0,0,0,0,0,0,0,0," & IIf(IsNull(lRS!plaa04), 0, lRS!plaa04) & ",0,0,0,0,"
   
   Sql = Sql & mcad & "0,0," & IIf(IsNull(lRS!plaapo), 0, lRS!plaapo) & "," & IIf(IsNull(lRS!pladed), 0, lRS!pladed) & "," & lRS!plaing & "," & lRS!planet & ",'" & Format(lRS!pladia, FormatFecha) & "','" & rs!CodAfp & "','T'," _
       & "" & IIf(IsNull(lRS!plad111), 0, lRS!plad111) & "," & IIf(IsNull(lRS!plad112), 0, lRS!plad112) & "," & IIf(IsNull(lRS!plad113), 0, lRS!plad113) & "," & IIf(IsNull(lRS!plad114), 0, lRS!plad114) & "," & IIf(IsNull(lRS!plad115), 0, lRS!plad115) & ",'" & mtrab & "','','','','','','',0,0,0,0,0,'" & rs!NUMAFP & "'," & rs!importe & ",'" & Format(lRS!pladia, FormatFecha) & "','','')"
   cn.Execute Sql
   rs.Close
   lRS.MoveNext
   I = I + 1
   Me.Caption = Str(I) & "  DE  " & Str(Barra.Max)
Loop

Sql$ = wFinTrans
cn.Execute Sql$
Screen.MousePointer = vbDefault
End Sub

Private Sub Command8_Click()
Dim I As Integer
Dim mcad As String
Screen.MousePointer = vbHourglass
Sql$ = wInicioTrans
cn.Execute Sql$
Dathistoric.Recordset.MoveFirst
Num = 1
I = 1

Set lRS = Datplamas.Recordset
Barra.Max = Datplamas.Recordset.RecordCount
Dim mbol As String
Dim mtrab As String
Do While (Not lRS.EOF)
   Barra.Value = I
   mcad = ""
   If Not IsNull(lRS!AFP) Then
      Select Case lRS!AFP
          Case Is = "10": mcad = "04"
          Case Is = "11": mcad = "02"
          Case Is = "12": mcad = "03"
          Case Is = "16": mcad = "01"
      End Select
   End If
   Sql = "insert into planillas values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "'," _
       & "'" & Trim(Left(lRS!planom, 15)) & "','" & Trim(Mid(lRS!planom, 16, 15)) & "','','" & Trim(Mid(lRS!planom, 31, 15)) & "'," _
       & "'','001','" & Format(lRS!planac, FormatFecha) & "','01','M','','" & Left(lRS!pladocs, 8) & "','','','" & Mid(lRS!pladocs, 19, 15) & "'," _
       & "'" & mcad & "','" & Trim(lRS!afpcod) & "',null,'','" & Trim(lRS!plaloc) & "','','" & Trim(lRS!pladom) & "','','','001150101','','','','','',"
   If lRS!plafab = "A" Then mcad = "'01'," Else mcad = "'02',"
   Sql = Sql & mcad
   
   If Left(lRS!PlaCod, 1) = "E" Then
      Select Case lRS!placos
             Case Is = "01": mcad = "06"
             Case Is = "02": mcad = "01"
             Case Is = "03": mcad = "04"
             Case Is = "04": mcad = "03"
             Case Is = "05": mcad = "02"
             Case Is = "06": mcad = "09"
             Case Is = "07": mcad = "10"
             Case Is = "08": mcad = "19"
             Case Is = "09": mcad = "12"
             Case Is = "10": mcad = "13"
             Case Is = "11": mcad = "14"
             Case Is = "12": mcad = "24"
             Case Is = "13": mcad = "''"
             Case Is = "14": mcad = "11"
             Case Is = "15": mcad = "08"
             Case Is = "16": mcad = "30"
      End Select
   Else
      Select Case lRS!placos
         Case Is = "01": mcad = "01"
         Case Is = "02": mcad = "''"
         Case Is = "03": mcad = "03"
         Case Is = "04": mcad = "19"
         Case Is = "05": mcad = "02"
         Case Is = "06": mcad = "15"
         Case Is = "07": mcad = "09"
         Case Is = "08": mcad = "04"
         Case Is = "10": mcad = "24"
         Case Is = "11": mcad = "''"
         Case Is = "12": mcad = "11"
         Case Is = "13": mcad = "28"
         Case Is = "14": mcad = "''"
         Case Is = "15": mcad = "08"
         Case Is = "16": mcad = "10"
      End Select
   End If
   Sql = Sql & mcad
   Sql = Sql & ",'" & Format(lRS!plaing, FormatFecha) & "',"
   If Left(lRS!PlaCod, 1) = "E" Then Sql = Sql & "'01',''," Else Sql = Sql & "'02','',"
   If IsNull(lRS!places) Then Sql = Sql & "null," Else Sql = Sql & "'" & Format(lRS!places, FormatFecha) & "',"
   Sql = Sql & "'" & lRS!plasind & "','02','','','S','" & lRS!ipssvida & "" & "',"
   
   If IsNull(lRS!banco) Then
      mcad = "'','',"
   Else
      If Trim(lRS!banco) = "" Then mcad = "'01',''," Else mcad = "'03','01',"
   End If
   
   Sql = Sql & mcad
   Sql = Sql & "'S/.','02','" & lRS!cta & "" & "','','','','','" & wuser & "'," & FechaSys & ",'" & wuser & "','',"
   If lRS!jubilado = "S" Then mcad = "" & FechaSys & "," Else mcad = "null,"
   
   Sql = Sql & mcad
   Sql = Sql & "'','S/.','','')"
   cn.Execute Sql
   
   Sql$ = "insert into maestroaux values('" & Format(I, "00000000") & "',''," _
        & "'" & Trim(Left(lRS!planom, 15)) & "','" & Trim(Mid(lRS!planom, 16, 15)) & "','','" & Trim(Mid(lRS!planom, 31, 15)) & "'," _
        & "'','','','X','')"
   cn.Execute Sql$
       
   If lRS!plabas <> 0 Then
       If Left(lRS!PlaCod, 1) = "E" Then mcad = "'04',240,'')" Else mcad = "'02',48,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','01','S/.'," & lRS!plabas & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   If lRS!plaesp <> 0 Then
       mcad = "'04',240,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','36','S/.'," & lRS!plaesp & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   If lRS!plafam <> 0 Then
       If Left(lRS!PlaCod, 1) = "E" Then mcad = "'04',240,'')" Else mcad = "'02',48,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','02','S/.'," & lRS!plafam & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   If lRS!platasa <> 0 Then
       mcad = "'01',8,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','08','S/.'," & lRS!platasa & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   If lRS!platser <> 0 Then
       mcad = "'04',240,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','04','S/.'," & lRS!platser & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   
   If lRS!plamov <> 0 Then
       mcad = "'01',8,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','03','S/.'," & lRS!plamov & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   If lRS!afp03 <> 0 Then
       If Left(lRS!PlaCod, 1) = "E" Then mcad = "'04',240,'')" Else mcad = "'02',48,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','05','S/.'," & lRS!afp03 & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   If lRS!afp10 <> 0 Then
       If Left(lRS!PlaCod, 1) = "E" Then mcad = "'04',240,'')" Else mcad = "'02',48,'')"
       Sql = "insert into plaremunbase values('01','" & lRS!PlaCod & "','" & Format(I, "00000000") & "','06','S/.'," & lRS!afp10 & ","
       Sql = Sql & mcad
       cn.Execute Sql$
   End If
   
   
   lRS.MoveNext
   I = I + 1
      Me.Caption = Str(I) & "  DE  " & Str(Barra.Max)
Loop

Sql$ = "Update numeracion set numero=" & I - 1 & " where documento='AUX'"
cn.Execute Sql$

Sql$ = wFinTrans
cn.Execute Sql$
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
'cncarga.ConnectionString = "driver={SQL server};" _
'& "server=SERVIDOR;" _
'& "uid=SA;" _
'& "pwd=NADASA;" _
'& "database=bdfacred"
'
'cncarga.ConnectionTimeout = 30
'cncarga.Open
End Sub
