VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRepafp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Reportes de Afp «"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "FrmRepafp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmRepafp.frx":030A
      Left            =   2250
      List            =   "FrmRepafp.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   900
      Width           =   2295
      Begin VB.OptionButton Opc2 
         Caption         =   "Lista"
         Height          =   195
         Left            =   1440
         TabIndex        =   18
         Top             =   200
         Width           =   735
      End
      Begin VB.OptionButton Opc1 
         Caption         =   "Formato"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   200
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.ComboBox Cmbtipotra 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1020
      Width           =   1935
   End
   Begin VB.CheckBox Chkdisco 
      Caption         =   "Remision con Medio Magnetico"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   2235
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox Cmbafp 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1830
      Width           =   4335
   End
   Begin VB.TextBox Txtano 
      Height          =   315
      Left            =   1335
      MaxLength       =   4
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      ItemData        =   "FrmRepafp.frx":039A
      Left            =   1320
      List            =   "FrmRepafp.frx":03AA
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1425
      Width           =   4335
   End
   Begin MSMask.MaskEdBox Txtfechapago 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2235
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   80
         Width           =   4335
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1020
      Width           =   1125
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AFP"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1830
      Width           =   300
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   315
      Left            =   1935
      TabIndex        =   10
      Top             =   600
      Width           =   255
      Size            =   "450;556"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Boletas"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1425
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Pago"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Peiodo"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "FrmRepafp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mmes As String
Dim vTipoTra As String
Dim VTipo As String
Dim vAfp As String
Dim mlinea As Integer

Private Sub Cmbafp_Click()
If Cmbafp.Text = "TOTAL" Then vAfp = "": mCodAfp = "" Else vAfp = fc_CodigoComboBox(Cmbafp, 2): mCodAfp = fc_CodigoComboBox(Cmbafp, 2)
End Sub

Private Sub Cmbcia_Click()
Cmbmes.ListIndex = Month(fecha) - 2
Txtano.Text = Format(Year(Date), "0000")
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipotra)
Cmbtipotra.AddItem "TOTAL"
Call fc_Descrip_Maestros2("01069", "", Cmbafp)
Cmbafp.AddItem "TOTAL"
End Sub

Private Sub Cmbmes_Click()
mmes = Format(Cmbmes.ListIndex + 1, "00")
End Sub

Private Sub CmbTipo_Click()
VTipo = Format(Cmbtipo.ListIndex + 1, "00")
End Sub

Private Sub Cmbtipotra_Click()
If Cmbtipotra.Text = "TOTAL" Then vTipoTra = "" Else vTipoTra = fc_CodigoComboBox(Cmbtipotra, 2)
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5865
Me.Height = 3075
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Cmbtipo.ListIndex = 0
End Sub
Public Sub Procesa_RepAfp()
If Opc1.Value = True Then
   If Not IsDate(Txtfechapago.Text) Then
      MsgBox "Debe Indicar la Fecha de Pago", vbInformation, "Formato de AFP"
      Exit Sub
   End If
   Procesa_RepAfp_Formato
Else
   Procesa_RepAfp_Lista
End If
End Sub
Private Sub Procesa_RepAfp_Formato()
Dim mCodAfp As String
Dim mNombArchivo As String
Dim mcad As String
Dim tot0 As Currency
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency
Dim totP1 As Currency
Dim totP2 As Currency
Dim totP3 As Currency
Dim totP4 As Currency
Dim totP5 As Currency
Dim mnumtra As Integer
Dim mnumemple As Integer
Dim minicio As Boolean
Dim I As Integer
Dim mtotret As Currency
Dim FILTRAAFP As String
Dim ultfila As Integer

minicio = True
mcad = ""
tot0 = 0: tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
totP1 = 0: totP2 = 0: totP3 = 0: totP4 = 0: totP5 = 0
mCodAfp = fc_CodigoComboBox(Cmbafp, 2)

Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='11' and status<>'*'"
'Debug.Print SQL$
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   Do While Not rs.EOF
      mcad = mcad & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
'   Debug.Print mcad
   mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
   If rs.State = 1 Then rs.Close
End If

'SQL = "SELECT a.CIA,a.PLACOD,a.FECHAPROCESO,a.CODAFP,b.numafp,b.codafp,b.tipotrabajador,c.d111 , c.d112, c.d113, c.d114, c.d115, c.remu"
'SQL = SQL & " FROM plahistorico a INNER JOIN planillas b ON (b.cia=a.cia and b.placod=a.placod and b.tipotrabajador like '" & Trim(vTipoTra) & "%' ) INNER JOIN (SELECT  PLACOD,sum(d111) as d111,   sum(d112) as d112,  sum(d113) as d113,  sum(d114) as d114,  sum(d115) as d115,"
'SQL = SQL & mcad & " FROM PLAHISTORICO WHERE cia='" & wcia & "' and status!='*' and MONTH(fechaproceso)=" & Val(mmes) & " and  YEAR(fechaproceso)=" & Val(Txtano.Text) & " AND PROCESO "& mcad &" GROUP BY PLACOD) c ON"
'SQL = SQL & " (c.placod=a.placod) WHERE a.cia='" & wcia & "' and a.status!='*' and MONTH(a.fechaproceso)=" & Val(mmes) & " and  YEAR(a.fechaproceso)=" & Val(Txtano.Text) & " and a.fechaproceso=(SELECT MAX(fechaproceso) "
'SQL = SQL & " FROM plahistorico WHERE cia=a.cia and status!='*' and MONTH(fechaproceso)=" & Val(mmes) & " and  YEAR(fechaproceso)=" & Val(Txtano.Text) & " and placod=a.placod)"
'SQL = SQL & " GROUP BY a.CIA,a.PLACOD,a.FECHAPROCESO,a.CODAFP,b.numafp,b.codafp,b.tipotrabajador,c.d111 , c.d112, c.d113, c.d114, c.d115, c.remu"
'SQL = SQL & " HAVING a.CODAFP='"& vAfp &"'"

If Cmbafp.Text = "TOTAL" Then
    FILTRAAFP = ""
Else
    FILTRAAFP = " AND DBO.fc_ultimoafp('" & wcia & "',a.PLACOD," & Val(Txtano.Text) & "," & Val(mmes) & ",0)=" & mCodAfp
End If
mcad = "sum" & mcad & " as remu,"
Sql$ = nombre()
Sql$ = Sql$ & " DBO.fc_ultimoafp('" & wcia & "',a.PLACOD," & Val(Txtano.Text) & "," & Val(mmes) & ",1) AS numafp, DBO.fc_ultimoafp('" & wcia & "',A.PLACOD," & Val(Txtano.Text) & "," & Val(mmes) & ",0) AS codafp,tipotrabajador, "
Sql$ = Sql$ & mcad
     
     'If VTipo = "01" Then mcad = "in ('01','03','04','05')" Else mcad = "='02'"
     If VTipo = "01" Then mcad = "in ('01','04','05')" Else mcad = "='02'"
     
     Sql$ = Sql$ & "sum(d111) as d111,sum(d112) as d112,sum(d113) as d113,sum(d114) as d114,sum(d115) as d115,fechavacai,b.fcese from plahistorico a,planillas b "
     Sql$ = Sql$ & " where a.cia='" & wcia & "' and proceso "
     Sql$ = Sql$ & mcad & " and year(fechaproceso)= " & Val(Txtano.Text) & " "
     Sql$ = Sql$ & "and month(fechaproceso)=" & Val(mmes) & " "
     Sql$ = Sql$ & "and  a.status<>'*' and b.cia=a.cia and a.placod=b.placod and b.status<>'*' and tipotrabajador like '" & Trim(vTipoTra) + "%" & "'"
     Sql$ = Sql$ & FILTRAAFP
     Sql$ = Sql$ & "and (b.txtrajub<>'S' or b.txtrajub is null) group by a.placod,b.ap_pat,b.ap_mat,b.nom_1,"
     Sql$ = Sql$ & "b.nom_2,b.numafp,b.codafp,b.tipotrabajador,"
     Sql$ = Sql$ & "a.fechavacai,B.FCESE order by b.codafp,a.placod"
     
If Not (fAbrRst(rs, Sql$)) Then
   If rs.State = 1 Then rs.Close
   MsgBox "No Hay registros a consultar", vbInformation, Me.Caption
   Exit Sub
End If
rs.MoveFirst
mlinea = 1
If Cmbafp.ListIndex = -1 Then
   mNombArchivo = "AFPTOTAL.txt"
Else
   mNombArchivo = "REPAFP" & Trim(Cmbafp.Text) & ".txt"
End If
RUTA$ = App.Path & "\REPORTS\" & mNombArchivo
Open RUTA$ For Output As #1
mCodAfp = ""
mnumtra = 0: mnumemple = 0
Do While Not rs.EOF
   mcad = ""
   If mCodAfp <> rs!CodAfp Then
      Cabeza_Afp (rs!CodAfp)
      mlinea = 1
      ultfila = 0
      mCodAfp = rs!CodAfp
   End If
   If mlinea >= 15 Then Cabeza_Afp (rs!CodAfp): mlinea = 1: ultfila = ultfila + 14
   mcad = mcad & " " & Format(mlinea + ultfila, "00") & " " & Chr(179) & " " & lentexto(15, Left(rs!NUMAFP, 15)) & " " & Chr(179) & lentexto(59, Left(rs!nombre, 59)) & Chr(179)
   If VTipo = "02" Then
        If Not IsNull(rs!fechavacai) Then
          mcad = mcad & " 05 " & Chr(179) & "  " & Format(rs!fechavacai, "dd/mm/yyyy") & "  " & Chr(179)
      Else
        If Not IsNull(rs!fcese) Then
            mcad = mcad & " 02 " & Chr(179) & "  " & Format(rs!fcese, "dd/mm/yyyy") & "  " & Chr(179)
        Else
            mcad = mcad & "    " & Chr(179) & "  " & "          " & "  " & Chr(179)
        End If
      End If
   Else
        If Not IsNull(rs!fcese) Then
            mcad = mcad & " 02 " & Chr(179) & "  " & Format(rs!fcese, "dd/mm/yyyy") & "  " & Chr(179)
        Else
            mcad = mcad & "    " & Chr(179) & "  " & "          " & "  " & Chr(179)
        End If
   End If
   mtotret = 0
   For I = 1 To 6
       Select Case I
              Case Is = 1
                   If rs(I + 3) = 0 Then
                      mcad = mcad & Space(14) & Chr(179)
                   Else
                      mcad = mcad & "  " & fCadNum(rs(I + 3), "#####,##0.00") & Chr(179)
                   End If
                   totP1 = totP1 + rs(I + 3)
              Case Is = 2
                   If rs(I + 3) = 0 Then
                      mcad = mcad & Space(13) & Chr(179)
                      mcad = mcad & Space(12) & Chr(179) & Space(12) & Chr(179) & Space(11) & Chr(179)
                      mcad = mcad & Space(13) & Chr(179)
                   Else
                      mcad = mcad & " " & fCadNum(rs(I + 3), "#####,##0.00") & Chr(179)
                      mcad = mcad & Space(12) & Chr(179) & Space(12) & Chr(179) & Space(11) & Chr(179)
                      mcad = mcad & " " & fCadNum(rs(I + 3), "#####,##0.00") & Chr(179)
                   End If
                   totP2 = totP2 + rs(I + 3)
              Case Is = 3
                   If rs(I + 3) = 0 Then
                      mcad = mcad & Space(11) & Chr(179)
                   Else
                      mcad = mcad & " " & fCadNum(rs(I + 3), "###,##0.00") & Chr(179)
                   End If
                   mtotret = mtotret + rs(I + 3)
                   totP3 = totP3 + rs(I + 3)
              Case Is = 5
                   If rs(I + 3) = 0 Then
                      mcad = mcad & Space(12) & Chr(179)
                   Else
                      mcad = mcad & " " & fCadNum(rs(I + 3), "####,##0.00") & Chr(179)
                   End If
                   mtotret = mtotret + rs(I + 3)
                   totP4 = totP4 + rs(I + 3)
      End Select
   Next I
   mcad = mcad & " " & fCadNum(mtotret, "##,###,##0.00") & Chr(179)
   totP5 = totP5 + mtotret
   'mcad = fCadNum(RS!remu, "###,##0.00") & "   " & fCadNum(RS!d111, "###,##0.00") & "    0.00  " & fCadNum(RS!d112, "###,##0.00") & " " & fCadNum(RS!d115, "###,##0.00") & "  " & fCadNum(RS!d114, "###,##0.00") & "   " & fCadNum(RS!d112 + RS!d113 + RS!d114 + RS!d115, "###,##0.00") & " " & fCadNum(RS!d111 + RS!d112 + RS!d113 + RS!d114 + RS!d115, "###,##0.00")
   Print #1, Chr(179) & mcad
   tot0 = tot0 + rs!remu: tot1 = tot1 + rs!d111: tot2 = tot2 + rs!d112: tot3 = tot3 + rs!d113: tot4 = tot4 + rs!d114: tot5 = tot5 + rs!d115
   mlinea = mlinea + 1
   rs.MoveNext
   If rs.EOF Then
      mcad = Linea("F")
      Print #1, mcad
      mcad = LineaTotPag(totP1, totP2, totP3, totP4, totP5)
      Print #1, mcad
      Call Final_Afp(tot0, tot1, tot2, tot4, tot2 + tot4)
      totP1 = 0: totP2 = 0: totP3 = 0: totP4 = 0: totP5 = 0
      tot0 = 0: tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
   Else
      If mCodAfp <> rs!CodAfp Or mlinea = 15 Then
         mcad = Linea("F")
         Print #1, mcad
         mcad = LineaTotPag(totP1, totP2, totP3, totP4, totP5)
         Print #1, mcad
         If mlinea = 15 Then
            Final_Pag
         Else
           Call Final_Afp(tot0, tot1, tot2, tot4, tot2 + tot4)
           tot0 = 0: tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
         End If
         totP1 = 0: totP2 = 0: totP3 = 0: totP4 = 0: totP5 = 0
      Else
         mcad = Linea("")
         Print #1, mcad
      End If
   End If
Loop
Close #1
Call Imprime_Txt(mNombArchivo, RUTA$)
End Sub
Private Sub Cabeza_Afp(CodAfp)
Dim mcad As String
Dim rs2 As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim wciamae As String
wciamae = Determina_Maestro("01069")
Sql$ = "Select * from maestros_2 where status<>'*' and cod_maestro2='" & CodAfp & "'"
Sql$ = Sql$ & wciamae
Print #1, Chr(27) + Chr(48) + Chr(27) + Chr(69)
If (fAbrRst(rs2, Sql$)) Then
   Print #1, Chr(18) & Chr(14) & "AFP " & rs2!DESCRIP & Chr(20)
End If
If rs2.State = 1 Then rs2.Close

Sql$ = "SELECT A.*,v.flag1 AS via,z.descrip as zona FROM CIA A LEFT OUTER JOIN maestros_2 z ON ( z.ciamaestro='01035' and  z.cod_maestro2= a.cod_zona)"
Sql$ = Sql$ & "    LEFT OUTER JOIN maestros_2 v ON ( v.ciamaestro='01036'  and v.cod_maestro2 =  a.cod_via) WHERE a.cod_cia='" & wcia & "'"
    
If (fAbrRst(rs2, Sql$)) Then
   Print #1, Space(110) & Chr(201) & String(20, Chr(205)) & Chr(187)
   Print #1, Space(110) & Chr(186) & Space(20) & Chr(186)
   Print #1, Space(56) & "PLANILLA DE PAGO DE APORTES PREVISIONALES  " & Space(11) & Chr(186) & Space(20) & Chr(186)
   Print #1, Space(110) & Chr(200) & String(20, Chr(205)) & Chr(188)
   Print #1, Space(25) & "PERIODO DE DEVENGUE :   " & mmes & Space(1) & Format(Txtano.Text, "0000") & Space(5) & "REMISION CON MEDIO MAGNETICO : SI [  ]    NO [X]" & Chr(15) + Chr(27) + Chr(70)
   Print #1, Space(82) & Chr(192) & String(5, Chr(196)) & Chr(193) & String(8, Chr(196)) & Chr(217)
   Print #1, Chr(218) & String(50, Chr(196)) & Chr(191)
   Print #1, Chr(179) & "SECCION I     IDENTIFICACION DEL EMPLEADOR        " & Chr(179)
   Print #1, Chr(195) & String(50, Chr(196)) & Chr(193) & String(39, Chr(196)) & Chr(194) & String(20, Chr(196)) & Chr(194) & String(29, Chr(196)) & Chr(194) & String(28, Chr(196)) & Chr(194) & String(52, Chr(196)) & Chr(191)
   Print #1, Chr(179) & "Nombre o Razon Social" & Space(69) & Chr(179) & "R.U.C.              " & Chr(179) & "Cta. Banca." & Chr(179) & "No. Cuenta       " & Chr(179) & "Tipo Cuenta                 " & Chr(179) & "Institucion Financiera                              " & Chr(179)
   mcad = lentexto(52, Left(rs2!razsoc, 52)) & " " & Left(rs2!RUC, 11) & "  " & lentexto(13, Left(rs2!afpnrocta, 13))
   mcad = mcad & "    " & IIf(rs2!afptipocta = "02", "CTA. AHORROS     ", "CTA. CORRIENTE   ")
   wciamae = Determina_Maestro("01007")
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs2!afpbanco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rsTemp, Sql$)) Then mcad = mcad & lentexto(30, Left(rsTemp!DESCRIP, 30)) & Chr(179)
   If rsTemp.State = 1 Then rsTemp.Close
   Print #1, Chr(179) & Chr(18) + Chr(27) + Chr(71) & mcad & Chr(13) + Chr(27) + Chr(70) + Chr(15)
   Print #1, Chr(195) & String(90, Chr(196)) & Chr(193) & String(5, Chr(196)) & Chr(194) & String(14, Chr(196)) & Chr(193) & String(29, Chr(196)) & Chr(193) & String(11, Chr(196)) & Chr(194) & String(16, Chr(196)) & Chr(193) & String(52, Chr(196)) & Chr(180)
   Print #1, Chr(179) & "Direccion" & Space(87) & Chr(179) & "N0 Departamento/Interior Manzana Lote                   " & Chr(179) & "Urbanizacion Localidad" & Space(47) & Chr(179)
   mcad = lentexto(56, Left(Trim(rs2!via) & " " & rs2!direcc, 56)) & " " & lentexto(24, Left(rs2!NRO & " " & rs2!DPTO, 24)) & Space(9)
   Sql$ = "select * from tablas where cod_tipo='01' and status<>'*' and cod_maestro2='" & rs2!urb & "'"
   If (fAbrRst(rsTemp, Sql$)) Then mcad = mcad & lentexto(38, Left(rsTemp!DESCRIP, 38)) & "  " & Chr(179)
   If rsTemp.State = 1 Then rsTemp.Close
   Print #1, Chr(179) & Chr(18) & mcad & " " & Chr(15)
   Print #1, Chr(195) & String(76, Chr(196)) & Chr(194) & String(19, Chr(196)) & Chr(193) & String(36, Chr(196)) & Chr(194) & String(19, Chr(196)) & Chr(193) & String(49, Chr(196)) & Chr(194) & String(19, Chr(196)) & Chr(180)
   Print #1, Chr(179) & "Distrito" & Space(68) & Chr(179) & "Provincia" & Space(47) & Chr(179) & "Departamento" & Space(57) & Chr(179) & "Telefono           " & Chr(179)
   Sql$ = "select dist,prov,dpto from dbo.fc_distrito_ubigeo_sunat( '" & rs2!cod_ubi & "')"
   If (fAbrRst(rsTemp, Sql$)) Then mcad = lentexto(40, Left(rsTemp!DIST, 40)) & "      " & lentexto(30, Left(rsTemp!PROV, 30)) & "   " & lentexto(38, Left(rsTemp!DPTO, 38)) & "   " Else Space (100)
   If rsTemp.State = 1 Then rsTemp.Close
   Sql$ = "select telef from telef_cia where cod_cia='" & wcia & "'"
   If (fAbrRst(rsTemp, Sql$)) Then mcad = mcad & lentexto(10, Left(rsTemp!telef, 10))
   If rsTemp.State = 1 Then rsTemp.Close
   Print #1, Chr(179) & Chr(18) & mcad & Chr(179) & Chr(15)
   Print #1, Chr(195) & String(20, Chr(196)) & Chr(194) & String(55, Chr(196)) & Chr(193) & String(11, Chr(196)) & Chr(194) & String(17, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(193) & Chr(194) & String(48, Chr(196)) & Chr(194) & String(19, Chr(196)) & Chr(193) & String(19, Chr(196)) & Chr(180)
   Print #1, Chr(179) & "Representante Legal " & Chr(179) & Space(67) & Chr(179) & "Tipo doc. Ident. " & Chr(179) & "No Documento Identificacion" & Chr(179) & "Elaborado por:Apellidos y Nombres" & Space(15) & Chr(179) & "Area o Departamento" & Space(5) & "Telefono       " & Chr(179)
   mcad = lentexto(52, Left(rs2!rep_nom, 52)) & " D.N.I.    " & lentexto(15, Left(IIf(IsNull(rs2!num_doc), "", rs2!num_doc), 15)) & " " & lentexto(27, Left(rs2!afpresponsable, 27)) & " "
   wciamae = Determina_Maestro("01044")
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs2!afparearesp & "'"
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rsTemp, Sql$)) Then mcad = mcad & lentexto(13, Left(rsTemp!DESCRIP, 13)) & " "
   If rsTemp.State = 1 Then rsTemp.Close
   mcad = mcad & lentexto(9, Left(rs2!afpresptlf, 9))
   Print #1, Chr(179) & Chr(18) & mcad & Chr(179) & Chr(15)
   Print #1, Chr(192) & String(88, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(193) & String(48, Chr(196)) & Chr(193) & String(39, Chr(196)) & Chr(217)
   
   Print #1, Chr(218) & String(82, Chr(196)) & Chr(194) & String(100, Chr(196)) & Chr(194) & String(39, Chr(196)) & Chr(191)
   Print #1, Chr(179) & "SECCION II DETALLE DE APORTE OBLIGATORIOS Y VOLUNTARIOS" & Space(27) & Chr(179) & Space(40) & "FONDO   DE  PENSIONES" & Space(39) & Chr(179) & "    RETENCIONES Y RETRIBUCIONES" & Space(8) & Chr(179)
   Print #1, Chr(195) & String(82, Chr(196)) & Chr(197) & String(19, Chr(196)) & Chr(194) & String(14, Chr(196)) & Chr(197) & String(13, Chr(196)) & Chr(194) & String(25, Chr(196)) & Chr(194) & String(11, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(197) & String(11, Chr(196)) & Chr(194) & String(12, Chr(196)) & Chr(194) & String(14, Chr(196)) & Chr(180)
   Print #1, Chr(179) & Space(30) & "IDENTIFICACION" & Space(38) & Chr(179) & "Movto. del Personal" & Chr(179) & " Remuneracion " & Chr(179) & "   Aporte    " & Chr(179) & "    Aporte Voluntario    " & Chr(179) & "   Aporte  " & Chr(179) & "    TOTAL    " & Chr(179) & "           " & Chr(179) & "    Comis.  " & Chr(179) & "     TOTAL    " & Chr(179)
   Print #1, Chr(195) & String(4, Chr(196)) & Chr(194) & String(17, Chr(196)) & Chr(194) & String(59, Chr(196)) & Chr(197) & String(4, Chr(196)) & Chr(194) & String(14, Chr(196)) & Chr(180) & "  Asegurable  " & Chr(179) & "Obligatorio. " & Chr(195) & String(12, Chr(196)) & Chr(194) & String(12, Chr(196)) & Chr(180) & "  Emplead. " & Chr(179) & "    FONDO    " & Chr(179) & "  Seguros  " & Chr(179) & "   Sobre RA " & Chr(179) & "   RETENCION  " & Chr(179)
   Print #1, Chr(179) & "Afil" & Chr(179) & "Codigo Unico SPP " & Chr(179) & Space(21) & "Apellidos y Nombre" & Space(20) & Chr(179) & "Tipo" & Chr(179) & "    Fecha     " & Chr(179) & Space(14) & Chr(179) & Space(13) & Chr(179) & "C/fin Prev. " & Chr(179) & "S/fin Prev. " & Chr(179) & Space(11) & Chr(179) & "  PENSIONES  " & Chr(179) & Space(11) & Chr(179) & "      S/.   " & Chr(179) & "  RETRIBUCION " & Chr(179)
   mcad = Linea("")
   Print #1, mcad
End If
If rs2.State = 1 Then rs2.Close
End Sub
Private Sub Procesa_RepAfp_Lista()
Dim mCodAfp As String
Dim mNombArchivo As String
Dim mcad As String
Dim tot0 As Currency
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency
Dim mnumtra As Integer
Dim mnumemple As Integer
Dim minicio As Boolean
Dim strConOnpAfp1 As String
Dim strConOnpAfp2 As String

'Adecuación para cuando el trabajador cuenta con aportaciones al ONP y AFP

'strConOnpAfp1 = "(CASE WHEN a.codafp = '01' AND b.codafp<>'01' THEN b.codafp ELSE a.codafp END)"
'strConOnpAfp2 = "a.d11<>(CASE WHEN a.codafp = '01' AND b.codafp<>'01' THEN 1 ELSE 0 END) "

strConOnpAfp1 = " b.codafp"
strConOnpAfp2 = "(b.txtrajub<>'S' or b.txtrajub is null)"

minicio = True
mcad = ""
tot0 = 0: tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='11' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   Do While Not rs.EOF
      mcad = mcad & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
   If rs.State = 1 Then rs.Close
End If
mcad = "sum" & mcad & " as remu,"
Sql$ = nombre()

Sql$ = Sql$ & "b.numafp," & strConOnpAfp1 & " AS codafp,tipotrabajador, "

'Sql$ = Sql$ & "b.numafp,b.codafp,tipotrabajador, "
     Sql$ = Sql$ & mcad
    
    'Hasta 2014 If VTipo = "01" Then mcad = "in ('01','03','04','05')" Else mcad = "='02'"
    
     'If VTipo = "04" Then mcad = "in ('01','02','03','04','05')" Else mcad = "='" & VTipo & "'"
     
     If VTipo = "04" Then mcad = "in ('01','02','04','05')" Else mcad = "='" & VTipo & "'"
     
     Sql$ = Sql$ & "sum(d111) as d111,sum(d112) as d112,sum(d113) as d113,sum(d114) as d114,sum(d115) as d115 from plahistorico a,planillas b " _
     & "where a.cia='" & wcia & "' and proceso "
     Sql$ = Sql$ & mcad & " and year(fechaproceso)= " & Val(Txtano.Text) & " " _
     & "and month(fechaproceso)=" & Val(mmes) & " and b.codafp like '" & Trim(vAfp) + "%" & "' " _
     & "and  a.status<>'*' and b.cia=a.cia and a.placod=b.placod and b.status<>'*' and tipotrabajador like '" & Trim(vTipoTra) + "%" & "' " _
     & "and " & strConOnpAfp2 & " group by a.placod,b.ap_pat,b.ap_mat,b.nom_1,b.nom_2,b.numafp," & strConOnpAfp1 & ",b.tipotrabajador " & _
     "order by " & strConOnpAfp1 & ",a.placod"
     
    '& "and  d11<>0 group by a.placod,b.ap_pat,b.ap_mat,b.nom_1,b.nom_2,b.numafp,b.codafp,b.tipotrabajador " & _

If Not (fAbrRst(rs, Sql$)) Then
   If rs.State = 1 Then rs.Close
   MsgBox "No hay registros a consultar", vbCritical, Me.Caption
   Exit Sub
End If
rs.MoveFirst
mlinea = 1
mcad = ""
If Cmbafp.ListIndex = -1 Then
   mNombArchivo = "REPAFPTOTAL.txt"
Else
   mNombArchivo = "REPAFP" & Trim(Cmbafp.Text) & ".txt"
End If
RUTA$ = App.Path & "\REPORTS\" & mNombArchivo
Open RUTA$ For Output As #1
mCodAfp = ""
mnumtra = 0: mnumemple = 0
Do While Not rs.EOF
   If mCodAfp <> rs!CodAfp Then
      mCodAfp = rs!CodAfp
      If minicio <> True Then
         Print #1,
         mcad = "T O T A L E S : .........S/. " & Space(15) & fCadNum(tot0, "###,##0.00") & "   " & fCadNum(tot1, "###,##0.00") & "    0.00  " & fCadNum(tot2, "###,##0.00") & " " & fCadNum(tot5, "###,##0.00") & "  " & fCadNum(tot4, "###,##0.00") & "   " & fCadNum(tot2 + tot3 + tot4 + tot5, "###,##0.00") & " " & fCadNum(tot1 + tot2 + tot3 + tot4 + tot5, "###,##0.00")
         Print #1, mcad
         Print #1,
         Print #1, "No Trabajadores           " & fCadNum(mnumtra, "###,##0.00")
         Print #1, "No Trabajadores Empleado  " & fCadNum(mnumemple, "###,##0.00")
         Print #1, "No Trabajadores Obrero    " & fCadNum(mnumtra - mnumemple, "###,##0.00")
         Print #1, Chr(12) + Chr(13)
      End If
      Cabeza_Afp_Lista (rs!CodAfp)
      tot0 = 0: tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
      mnumtra = 0: mnumemple = 0
   End If
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Cabeza_Afp_Lista (rs!CodAfp): mlinea = 1
   mcad = lentexto(15, Left(rs!NUMAFP, 15)) & lentexto(27, Left(rs!nombre, 27)) & Space(2) & fCadNum(rs!remu, "###,##0.00") & "   " & fCadNum(rs!d111, "###,##0.00") & "    0.00  " & fCadNum(rs!d112, "###,##0.00") & " " & fCadNum(rs!d115, "###,##0.00") & "  " & fCadNum(rs!d114, "###,##0.00") & "   " & fCadNum(rs!d112 + rs!d113 + rs!d114 + rs!d115, "###,##0.00") & " " & fCadNum(rs!d111 + rs!d112 + rs!d113 + rs!d114 + rs!d115, "###,##0.00")
   Print #1, mcad
   tot0 = tot0 + rs!remu: tot1 = tot1 + rs!d111: tot2 = tot2 + rs!d112: tot3 = tot3 + rs!d113: tot4 = tot4 + rs!d114: tot5 = tot5 + rs!d115
   mnumtra = mnumtra + 1
   If rs!TipoTrabajador = "01" Then mnumemple = mnumemple + 1
   mlinea = mlinea + 1
   minicio = False
   rs.MoveNext
Loop
Print #1,
mcad = "T O T A L E S : .........S/. " & Space(15) & fCadNum(tot0, "###,##0.00") & "   " & fCadNum(tot1, "###,##0.00") & "    0.00  " & fCadNum(tot2, "###,##0.00") & " " & fCadNum(tot5, "###,##0.00") & "  " & fCadNum(tot4, "###,##0.00") & "   " & fCadNum(tot2 + tot3 + tot4 + tot5, "###,##0.00") & " " & fCadNum(tot1 + tot2 + tot3 + tot4 + tot5, "###,##0.00")
Print #1, mcad
Print #1,
Print #1, "No Trabajadores           " & fCadNum(mnumtra, "###,##0.00")
Print #1, "No Trabajadores Empleado  " & fCadNum(mnumemple, "###,##0.00")
Print #1, "No Trabajadores Obrero    " & fCadNum(mnumtra - mnumemple, "###,##0.00")
Close #1
Call Imprime_Txt(mNombArchivo, RUTA$)
End Sub
Private Sub Cabeza_Afp_Lista(CodAfp)
Dim mcad As String
Dim rs2 As ADODB.Recordset
Dim mperiodo As String
Dim mdesafp As String
Dim mdescia As String
mdesafp = ""
mperiodo = Format(Txtano.Text, "0000") & Format(mmes, "00")
Sql$ = "Select * from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs2, Sql$)) Then mdescia = lentexto(40, Left(rs2!razsoc, 40))
If rs2.State = 1 Then rs2.Close
Sql$ = "select * from plaafp where codafp='" & CodAfp & "' and periodo='" & mperiodo & "' and status<>'*'"
If (fAbrRst(rs2, Sql$)) Then
   mdesafp = lentexto(36, Left(rs2!Descripcion, 36))
   Print #1, Chr(18) & mdescia & Space(20) & Format(Date, "dd/mm/yyyy") & Chr(15)
   Print #1,
   Print #1, Space(25) & " I N F O R M A C I O N   P A R A   D E S C U E N T O S   A   F A V O R   D E   A. F. P. "
   Print #1, Space(53) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000")
   Print #1, String(137, "-")
   Print #1, "   CODIGO              NOMBRE                  RESUMEN       FONDO    SOLIDAR.   PRIMA DE    COMISION  COMISION     SUBTOTAL     TOTAL"
   Print #1, "                                               ASIGNAC.      PROPIO   I.P.S.S.    SEGURO       FIJA    %SOBRE RA   APORTES AFP   GENERAL"
   Print #1, String(137, "-")
   Print #1, Space(18) & mdesafp & "   %  " & fCadNum(rs2!afp01, "##0.000") & "      0.000  " & fCadNum(rs2!afp02, "##0.000") & "   S/." & fCadNum(rs2!AFP05, "##0.00") & "     " & fCadNum(rs2!afp04, "##0.00")
   Print #1, String(137, "-")
End If
mlinea = 10
If rs2.State = 1 Then rs2.Close
End Sub
Private Function Linea(m As String) As String
If m = "F" Then
   Linea = Chr(192) & String(4, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(59, Chr(196)) & Chr(193) & String(4, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(197) & String(13, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(197) & String(11, Chr(196)) & Chr(197) & String(13, Chr(196)) & Chr(197) & String(11, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(180)
Else
   Linea = Chr(195) & String(4, Chr(196)) & Chr(197) & String(17, Chr(196)) & Chr(197) & String(59, Chr(196)) & Chr(197) & String(4, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(197) & String(13, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(197) & String(11, Chr(196)) & Chr(197) & String(13, Chr(196)) & Chr(197) & String(11, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(180)
End If
End Function
Private Function LineaTotPag(t1 As Currency, t2 As Currency, t3 As Currency, t4 As Currency, t5 As Currency) As String
Dim mcad As String
mcad = fCadNum(t2, "##,###,##0.00") & Chr(179)
mcad = mcad & Space(12) & Chr(179) & Space(12) & Chr(179) & Space(11) & Chr(179)
mcad = mcad & fCadNum(t2, "##,###,##0.00") & Chr(179)
LineaTotPag = Space(88) & Chr(179) & "Total Pagina  " & Chr(179) & fCadNum(t1, "###,###,##0.00") & Chr(179) & mcad & fCadNum(t3, "####,##0.00") & Chr(179) & fCadNum(t4, "#,###,##0.00") & Chr(179) & fCadNum(t5, "###,###,##0.00") & Chr(179)
End Function
Private Sub Final_Pag()
Dim mm As String
mm = Name_Month(Mid(Txtfechapago.Text, 4, 2))
Print #1, Chr(201) & String(85, Chr(205)) & Chr(187) & " " & Chr(192) & String(14, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(13, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(193) & String(11, Chr(196)) & Chr(193) & String(13, Chr(196)) & Chr(193) & String(11, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(217)
Print #1, Chr(186) & "  DECLARO BAJO JURAMENTO QUE LOS DATOS CONSIGNADOS SON EXPRESION FIEL DE LA REALIDAD " & Chr(186)
Print #1, Chr(186) & Space(85) & Chr(186)
Print #1, Chr(186) & Space(85) & Chr(186) & Space(94) & "Fecha de Pago :  " & Mid(Txtfechapago.Text, 1, 2) & " de " & mm & " de " & Mid(Txtfechapago.Text, 7, 4)
mm = lentexto(65, Left(String(47, Chr(196)), 65))
Print #1, Chr(186) & Space(20) & mm & Chr(186)
mm = lentexto(61, Left("Firma del Empleador o Representante Legal", 61))
Print #1, Chr(186) & Space(24) & mm & Chr(186)
Print #1, Chr(200) & String(85, Chr(205)) & Chr(188)
Print #1, Chr(12) + Chr(13)
End Sub
Private Sub Final_Afp(t1 As Currency, t2 As Currency, t3 As Currency, t4 As Currency, t5 As Currency)
Dim mm As String
mm = Name_Month(Mid(Txtfechapago.Text, 4, 2))
Print #1, Chr(201) & String(85, Chr(205)) & Chr(187) & " " & Chr(192) & String(14, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(13, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(193) & String(11, Chr(196)) & Chr(193) & String(13, Chr(196)) & Chr(193) & String(11, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(217)
mcad = " " & Chr(201) & String(14, Chr(205)) & Chr(203) & String(14, Chr(205)) & Chr(203) & String(13, Chr(205)) & Chr(203) & String(12, Chr(205)) & Chr(203) & String(12, Chr(205)) & Chr(203) & String(11, Chr(205)) & Chr(203) & String(13, Chr(205)) & Chr(203) & String(11, Chr(205)) & Chr(203) & String(12, Chr(205)) & Chr(203) & String(14, Chr(205)) & Chr(187)
Print #1, Chr(186) & "  DECLARO BAJO JURAMENTO QUE LOS DATOS CONSIGNADOS SON EXPRESION FIEL DE LA REALIDAD " & Chr(186) & mcad
mcad = fCadNum(t2, "##,###,##0.00") & Chr(186)
mcad = mcad & Space(12) & Chr(186) & Space(12) & Chr(186) & Space(11) & Chr(186)
mcad = mcad & fCadNum(t2, "##,###,##0.00") & Chr(186)
mcad = Chr(186) & Space(85) & Chr(186) & " " & Chr(186) & "Total General " & Chr(186) & fCadNum(t1, "###,###,##0.00") & Chr(186) & mcad & fCadNum(t3, "####,##0.00") & Chr(186) & fCadNum(t4, "#,###,##0.00") & Chr(186) & fCadNum(t5, "###,###,##0.00") & Chr(186)
Print #1, mcad
mcad = Chr(186) & Space(85) & Chr(186) & " " & Chr(200) & String(14, Chr(205)) & Chr(202) & String(14, Chr(205)) & Chr(202) & String(13, Chr(205)) & Chr(202) & String(12, Chr(205)) & Chr(202) & String(12, Chr(205)) & Chr(202) & String(11, Chr(205)) & Chr(206) & String(13, Chr(205)) & Chr(206) & String(11, Chr(205)) & Chr(202) & String(12, Chr(205)) & Chr(206) & String(14, Chr(205)) & Chr(185)
Print #1, mcad
Print #1, Chr(186) & Space(85) & Chr(186) & Space(62) & "Intereses Moratorios " & Chr(186) & Space(13) & Chr(186) & "  Intereses Moratorios  " & Chr(186) & Space(14) & Chr(186)
Print #1, Chr(186) & Space(85) & Chr(186) & Space(83) & Chr(204) & String(13, Chr(205)) & Chr(185) & "                        " & Chr(204) & String(14, Chr(205)) & Chr(185)
mm = lentexto(65, Left(String(47, Chr(196)), 65))
mcad = Chr(186) & Space(20) & mm & Chr(186) & Space(58) & "Total Fondo de Pensiones " & Chr(186) & fCadNum(t2, "##,###,##0.00") & Chr(186) & " Total Reten. y Retrib. " & Chr(186) & fCadNum(t5, "###,###,##0.00") & Chr(186)
Print #1, mcad
mm = lentexto(61, Left("Firma del Empleador o Representante Legal", 61))
mcad = Chr(186) & Space(24) & mm & Chr(186) & Space(83) & Chr(200) & String(13, Chr(205)) & Chr(188) & "                        " & Chr(200) & String(14, Chr(205)) & Chr(188)
Print #1, mcad
mcad = Chr(200) & String(85, Chr(205)) & Chr(188)
Print #1, mcad
mm = Name_Month(Mid(Txtfechapago.Text, 4, 2))
Print #1, Space(87) & Space(94) & "Fecha de Pago :  " & Mid(Txtfechapago.Text, 1, 2) & " de " & mm & " de " & Mid(Txtfechapago.Text, 7, 4)
Print #1, Chr(12) + Chr(13)
End Sub

Private Sub SpinButton1_SpinDown()
If Txtano.Text = "" Then Txtano.Text = "0"
If Txtano.Text > 0 Then Txtano = Txtano - 1

End Sub

Private Sub SpinButton1_SpinUp()
If Txtano.Text = "" Then Txtano.Text = "0"
Txtano = Txtano + 1
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub

