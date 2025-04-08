VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCertretecionApf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Certeficado de AFP"
   ClientHeight    =   2670
   ClientLeft      =   2760
   ClientTop       =   4020
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BTN_REPORTE 
      Caption         =   "&Ver Reporte"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton CmdFind 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   400
      End
      Begin VB.TextBox TxtCodTrab 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox Cmbafp 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Txtano 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label LblNomTrab 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Trabajador"
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S.Pensionario:"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   1245
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   255
         Size            =   "450;564"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   405
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
         TabIndex        =   3
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Data dat 
      Caption         =   "Dta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   885
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCertretecionApf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlinea As Integer
Dim vAfp As String
Dim mciadir As String
Dim mciatlf As String
Dim rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim VConcepto As String
Private Sub Cmbafp_Click()
vAfp = fc_CodigoComboBox(Cmbafp, 2)
If vAfp = "01" Or vAfp = "02" Then
    VConcepto = "04"
Else
    VConcepto = "11"
End If
procesa_certificado
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
Call fc_Descrip_Maestros2("01069", "", Cmbafp)
End Sub

Private Sub CmdFind_Click()
Unload Frmgrdpla
Load Frmgrdpla
Frmgrdpla.Show vbModal
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 6120
Me.Height = 3105 '1605
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
Txtano.Text = Format(Year(Date), "0000")
Crea_Tablas
End Sub
Sub procesa_certificado()
Dim mItem As Integer
Dim mcad As String
Dim mcadIA As String
Dim mcadID As String
Dim cadnombre As String
Dim mtot As Currency
Dim mtipoB As String
Dim MUIT As Currency
Dim mtope As Currency
Dim mFactor As Currency
Dim t1 As Currency: Dim t2 As Currency: Dim t3 As Currency: Dim t4 As Currency
Dim t5 As Currency: Dim t6 As Currency: Dim t7 As Currency: Dim t8 As Currency
Dim flg_activo As String
Dim sc_activo As String
cadnombre = nombre()
mSemana = ""
mpag = 0
mtipoB = "D"
'VConcepto = "11"
MUIT = 0
flg_activo = "1"
sc_activo = ""
If flg_activo = "1" Then sc_activo = " and (fcese is null or fcese>'01/01/2019') "

Sql$ = "select uit from plauit where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then MUIT = rs!uit
If rs.State = 1 Then rs.Close

mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Normal
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"

If VConcepto = "04" Then
   Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' AND codafp='" & vAfp & "' " & " Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1

Do While Not rs.EOF
   Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'" & sc_activo
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PlaCod = rs!PlaCod
   
'   Sql$ = "SELECT PLACOD,CASE  SUM(D11) WHEN 0 THEN SUM(D04) " & _
'        " ELSE SUM(D11) END AS APORTE " & _
'        " FROM plahistorico  " & _
'        " WHERE proceso<>'05' " & _
'        " AND placod ='" & Rs!PLACOD & "' " & _
'        " AND codafp='" & vAfp & "' and year(fechaproceso)=" & Val(Txtano.Text) & _
'        " AND STATUS<>'*'" & _
'        " GROUP BY PLACOD"
'
'   If Not (fAbrRst(RS3, Sql$)) Then Else RS3.MoveFirst

   dat.Recordset!nombre = mcad
   If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
   dat.Recordset!importe = rs!APO
   dat.Recordset!AFP = rs!AFP
   'dat.Recordset!util = Rs!util
   dat.Recordset!util = 0
   dat.Recordset!afp03 = rs!afp03
   If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = " "
   dat.Recordset.Update
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"

'Ingresos Afectos para Aportacion Vacaciones
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"

If VConcepto = "04" Then
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' AND codafp='01'  Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' AND codafp='" & vAfp & "'  Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod, " & _
      "CASE CONVERT(VARCHAR(10),ISNULL(fcese,''),103)WHEN '01/01/1900' THEN '  /  /    ' ELSE CONVERT(VARCHAR(10),ISNULL(fcese,''),103) END AS FCESE " & _
      "from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'" & sc_activo
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = rs!PlaCod
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset!afp03 = rs!afp03
      'dat.Recordset!util = Rs!util
      dat.Recordset!util = 0
      dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy")
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then '
      dat.Recordset!sueldo = dat.Recordset!sueldo
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + rs!afp03
      'dat.Recordset!util = dat.Recordset!util + Rs!util
      dat.Recordset!util = 0
      dat.Recordset!vaca = dat.Recordset!vaca + rs!afectoa
      dat.Recordset.Update
      End If
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Gratificacion
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"

If VConcepto = "04" Then
   Sql$ = "Select placod,tipotrab, " & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05'  AND codafp='01' AND D04<>0 Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab, " & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05'  AND codafp='" & vAfp & "' AND D11<>0 Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'" & sc_activo
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = rs!PlaCod
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset!afp03 = rs!afp03
      'dat.Recordset!util = Rs!util
      dat.Recordset!util = 0
      If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = Null
      dat.Recordset.Update
   Else
   '***************************************************************
   ' MODIFICA   : 10/01/2008
   ' MOTIVO     : ERROR EN LA ASIGNACION DE SUMA SE COMENTA EL CODIGO
   ' MODIFICADO POR : RICARDO HINOSTROZA
   '****************************************************************
      dat.Recordset.Edit
     If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + rs!afp03
      'dat.Recordset!util = dat.Recordset!util ' + rs!AFP
      dat.Recordset!util = 0
      dat.Recordset.Update
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i18) as util "

If VConcepto = "04" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND codafp='01'  Group by placod,tipotrab order by placod" ','05'
Else
    Sql$ = "Select placod,tipotrab,sum(case when fechaproceso<convert(datetime,'02/05/2009',103) then i39+i40+I42+I43+I16 else i39+I42+I43+I16 end) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND codafp='" & vAfp & "'  Group by placod,tipotrab order by placod" ','05'
End If
If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'" & sc_activo
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = rs!PlaCod
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      'dat.Recordset!util = Rs!util
      dat.Recordset!util = 0
      If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = Null
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
       dat.Recordset!liquid = dat.Recordset!liquid + rs!afectoa
      'dat.Recordset!util = dat.Recordset!util + Rs!util
      dat.Recordset!util = 0
      dat.Recordset.Update
      
   End If
   rs.MoveNext
Loop
Dim IREM_NETA As Double

IREM_NETA = 0

Dim SISPENSION As String
SISPENSION = Cmbafp.Text

If Mid(SISPENSION, 1, 2) = "(L" Then
    SISPENSION = Space(10)
End If



If rs.State = 1 Then rs.Close

If Trim(Me.TxtCodTrab.Text) <> "" Then
    RUTA$ = App.Path & "\REPORTS\" & "QtaAnual_" & Trim(Me.TxtCodTrab.Text) & ".txt"
Else
    RUTA$ = App.Path & "\REPORTS\" & "QtaAnual.txt"
End If
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst

If Trim(Me.TxtCodTrab.Text) <> "" Then
    .FindFirst " placod='" & Trim(Me.TxtCodTrab.Text) & "'"
End If

miten = 1
mtot = 0
t1 = 0: t2 = 0: t3 = 0: t4 = 0
t5 = 0: t6 = 0: t7 = 0: t8 = 0
IREM_NETA = 0
Do While Not .EOF

   '***************************************************************
   ' MODIFICA   : 10/01/2008
   ' MOTIVO     : ERROR EN LA RESTA DEL TOPE DE SUMA SE COMENTA EL CODIGO
   ' MODIFICADO POR : RICARDO HINOSTROZA
   '****************************************************************
   
   
   mtope = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) - Val(!afp03)) - ((MUIT * 7))
   IREM_NETA = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util)) - Val(!afp03)
   
   mFactor = 0
   Select Case mtope
          Case Is < (Round(MUIT * 27, 2) + 1)
               mFactor = Round((mtope) * 0.15, 2)  '- (muit * 7)
          Case Is < (Round(MUIT * 54, 2) + 1)
               mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
          Case Else
               mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
    End Select
    If mFactor < 0 Then mFactor = 0
    Print #1, Chr(218) & String(46, Chr(196)) & Chr(194) & String(31, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "COMPROBANTE DE RETENCIONES POR APORTE         " & Chr(179) & "      EJERCICIO GRAVABLE       " & Chr(179)
    Print #1, Chr(179) & "AL SISTEMA DE PENSIONES                       " & Chr(179) & Space(31) & Chr(179)
    Print #1, Chr(179) & "                                              " & Chr(179) & Space(13) & Txtano.Text & Space(14) & Chr(179)
    Print #1, Chr(218) & String(78, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "Razon Social del Empleador: " & lentexto(32, Left(CmbCia.Text, 32)) & "   RUC " & wruc & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(34) & "CERTIFICA" & Space(35) & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "Que al Sr(a) : " & lentexto(50, Left(!nombre, 50)) & " CODIGO " & RTrim(!PlaCod) & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "se le ha retenido por aporte al sistema de pensiones el importe de(S/ " & Trim(fCadNum(!importe, "##,###.00")) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(194) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "1.- REMUNERACIONES" & Space(47) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Sueldo o Jornal Basico" & Space(39) & Chr(179) & fCadNum(!sueldo, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Incremento AFP" & Space(47) & Chr(179) & fCadNum(!AFP, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Gratificaciones" & Space(46) & Chr(179) & fCadNum(!grati, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Vacaciones     " & Space(46) & Chr(179) & fCadNum(!vaca, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Liquid/Subs    " & Space(46) & Chr(179) & fCadNum(!liquid, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "REMUNERACION TOTAL PARA RETENCIONES " & Space(29) & Chr(179) & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) + Val(!vaca) + Val(!liquid), "#,###,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "2.- IMPORTE RETENIDO" & Space(45) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    If Trim(SISPENSION) = "" Then
    Print #1, Chr(179) & "    ADMINISTRADORA DE FONDO DE PENSIONES AFP " & SISPENSION & Space(11) & Chr(179) & fCadNum(0, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    SISTEMA NACIONAL DEL PENSIONES "; Space(30) & Chr(179) & fCadNum(!importe, "#,###,###.00") & Chr(179)
    Else
    Print #1, Chr(179) & "    ADMINISTRADORA DE FONDO DE PENSIONES AFP " & SISPENSION & Space(11) & Chr(179) & fCadNum(!importe, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    SISTEMA NACIONAL DEL PENSIONES " & Space(30) & Chr(179) & Space(12) & Chr(179)
    End If
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "6.TOTAL RETENIDO Y APORTADO AL SISTEMA DE PENSIONES" & Space(14) & Chr(179) & fCadNum(!importe, "#,###,###.00") & Chr(179)
    Print #1, Chr(192) & String(65, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(217)
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1,
    Print #1, Space(36) & String(25, Chr(196)) & Space(2) & String(15, Chr(196))
    Print #1, Space(36) & "FIRMA DEL REPRESENTANTE" & Space(7) & "RECIBIDO"
    Print #1, Space(46) & "LEGAL"
    Print #1,
  '  Print #1, "  * APLICABLES SOLO PARA TRABAJADORES CON RENTAS DE OTRAS CATEGORIA"
    Print #1, SaltaPag
    
    If Trim(Me.TxtCodTrab.Text) <> "" Then
        Close #1
        Me.Cmbafp.ListIndex = -1
        Exit Do
        
    End If
   .MoveNext
Loop
End With
Close #1
Call Imprime_Txt("QtaAnual.txt", RUTA$)
dat.Database.Execute "delete from Tmpdeduc"
End Sub
Private Sub Crea_Tablas()

Dim tdf0 As TableDef
dat.DatabaseName = nom_BD
dat.Refresh
Set tdf0 = dat.Database.CreateTableDef("Tmpdeduc")
With tdf0
        .Fields.Append .CreateField("placod", dbText)
        .Fields(0).AllowZeroLength = True
        .Fields.Append .CreateField("nombre", dbText)
        .Fields(1).AllowZeroLength = True
        .Fields.Append .CreateField("sueldo", dbText)
        .Fields(2).AllowZeroLength = True
        .Fields(2).DefaultValue = "0.00"
        .Fields.Append .CreateField("afp", dbText)
        .Fields(3).AllowZeroLength = True
        .Fields(3).DefaultValue = "0.00"
        .Fields.Append .CreateField("grati", dbText)
        .Fields(4).AllowZeroLength = True
        .Fields(4).DefaultValue = "0.00"
        .Fields.Append .CreateField("vaca", dbText)
        .Fields(5).AllowZeroLength = True
        .Fields(5).DefaultValue = "0.00"
        .Fields.Append .CreateField("liquid", dbText)
        .Fields(6).AllowZeroLength = True
        .Fields(6).DefaultValue = "0.00"
        .Fields.Append .CreateField("total", dbText)
        .Fields(7).AllowZeroLength = True
        .Fields(7).DefaultValue = "0.00"
        .Fields.Append .CreateField("importe", dbText)
        .Fields(8).AllowZeroLength = True
        .Fields(8).DefaultValue = "0.00"
        .Fields.Append .CreateField("util", dbText)
        .Fields(9).AllowZeroLength = True
        .Fields(9).DefaultValue = "0.00"
        .Fields.Append .CreateField("afp03", dbText)
        .Fields(10).AllowZeroLength = True
        .Fields(10).DefaultValue = "0.00"
        .Fields.Append .CreateField("cese", dbText)
        .Fields(10).AllowZeroLength = True
       ' dat.Database.Execute "delete from Tmpdeduc"
        dat.Database.TableDefs.Append tdf0
    End With
    dat.RecordSource = "Tmpdeduc"
    dat.Refresh
    Set tdf0 = Nothing
End Sub

Private Sub Label4_Click()

End Sub

Private Sub SpinButton1_SpinDown()
If Val(Txtano.Text) > 0 Then Txtano = Txtano - 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
dat.RecordSource = ""
dat.Refresh
dat.Database.TableDefs.Delete "Tmpdeduc"
dat.Database.Close
End Sub

Private Sub TxtCodTrab_Change()
LblNomTrab.Caption = ""
End Sub

Private Sub TxtCodTrab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Cmbafp.SetFocus
End Sub

Private Sub Txtcodtrab_LostFocus()
If Trim(TxtCodTrab.Text) = "" Then Exit Sub
TxtCodTrab.Text = UCase(TxtCodTrab.Text)
Dim Sql As String
Dim xDni As String
Dim xCodigo As String
xDni = "": xCodigo = ""
'If Trim(TxtNro.Text) <> "" Then
'    If OptDNI(0).Value = True Then xCodigo = " and nro_doc='" & Trim(TxtNro.Text) & "' "
'    If OptCodigo(1).Value = True Then xCodigo = " and placod='" & Trim(TxtNro.Text) & "' "
'End If
xCodigo = " and placod='" & Trim(TxtCodTrab.Text) & "' "
Sql = "select dbo.fc_Razsoc_Trabajador(placod,cia) as nom_trab,fcese,PLACOD,codafp "
Sql = Sql & " from planillas where cia='" & wcia & "'" & xCodigo & " AND status<>'*'"
Dim Rq As ADODB.Recordset
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    If Not IsNull(Rq!fcese) Then
'        MsgBox "El Trabajador ya fue Cesado", vbExclamation, Me.Caption
'        GoTo Salir:
    End If
    LblNomTrab.Caption = Rq(0)
    Call rUbiIndCmbBox(Me.Cmbafp, Trim(Rq!CodAfp), "00")
'    LblPlacod.Caption = Trim(Rq!PlaCod & "")
Else
    LblNomTrab.Caption = ""
    MsgBox "No existe codigo,verifique", vbExclamation, Me.Caption
End If
Salir:
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
End Sub
