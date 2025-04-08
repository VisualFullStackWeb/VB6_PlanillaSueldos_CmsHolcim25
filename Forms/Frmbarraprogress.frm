VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmbarraprogress 
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Frmbarraprogress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Txtano 
         Height          =   285
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "Frmbarraprogress.frx":030A
         Left            =   600
         List            =   "Frmbarraprogress.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   820
         Visible         =   0   'False
         Width           =   2535
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   80
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   320
         Left            =   4320
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   255
         Size            =   "450;564"
      End
      Begin VB.Label Lblano 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Lblmes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         Top             =   840
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "Frmbarraprogress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlinea As Integer
Dim mpag As Integer
Dim mnomcia As String
Private Sub Form_Load()
Me.Top = 3200
Me.Left = 3500
Me.Height = 1245
Me.Width = 4800
Select Case NameForm
       Case Is = "TRABAJADORES"
            Me.Height = 1245
            lblmes.Visible = False
            Cmbmes.Visible = False
            Lblano.Visible = False
            Txtano.Visible = False
            SpinButton1.Visible = False
            Me.Caption = "Exportando Trabajadores"
       Case Is = "REMUNERACIONES"
            Me.Height = 1695
            lblmes.Visible = True
            Cmbmes.Visible = True
            Lblano.Visible = True
            Txtano.Visible = True
            SpinButton1.Visible = True
            Me.Caption = "Exportando Remuneraciones"
            Cmbmes.ListIndex = Month(Date) - 1
            Txtano.Text = Format(Year(Date), "0000")
End Select
End Sub
Private Sub Exporta_Trabajadores()

Dim rs2 As ADODB.Recordset
Dim ms As String
Dim td As String
Dim cad As String
Dim mcad As String
ms = "|"

Sql$ = " select ap_pat,ap_mat,nom_1,nom_2,tipo_doc,nro_doc,ruc_intermedio,placod,dni,lmilitar,pasaporte,fnacimiento,sexo,fingreso,tipotrabajador,fcese,ruc,essaludvida,codafp,sctr,afpfechaafil, " _
     & "nomvia,nrokmmza,intdptolote,nomzona,tvia,tzona,ubigeo from planillas where cia='" & wcia & "' and cat_trab<>'04' and status<>'*' and (fcese IS NULL OR fcese>'" & Format("01/" & Month(DateAdd("m", -1, Date)) & "/" & Year(DateAdd("m", -1, Date)), "MM/DD/YYYY") & "' )  "

If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Se encuentran datos para exportar ": Exit Sub

ProgressBar1.Min = 0
ProgressBar1.Max = rs.RecordCount
rs.MoveFirst
RUTA$ = App.Path & "\REPORTS\" & wruc & ".ase"
Open RUTA$ For Output As #1
Do While Not rs.EOF
   td = "": cad = "": mcad = ""
    'If Trim(rs!PLACOD) = "JO013" Then Stop
    td = rs!tipo_doc
    cad = rs!nro_doc
   cad = lentexto(15, Left(cad, 15))
   mcad = mcad & td & ms & cad & ms
   cad = lentexto(20, Left(rs!ap_pat, 20))
   mcad = mcad & cad & ms
   cad = lentexto(20, Left(rs!ap_mat, 20))
   mcad = mcad & cad & ms
   cad = lentexto(20, Left(rs!nom_1 & Space(1) & rs!nom_2, 20))
   mcad = mcad & cad & ms
   If IsNull(rs!fnacimiento) Then cad = Space(10) Else cad = Format(rs!fnacimiento, "dd/mm/yyyy")
   mcad = mcad & cad & ms
   If rs!sexo = "M" Then cad = "1" Else cad = "2"
   mcad = mcad & cad & ms
   Sql$ = "select * from platelefono where status<>'*' and cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then cad = Left(rs2!telefono, 7) Else cad = Space(7)
   If rs2.State = 1 Then rs2.Close
   mcad = mcad & cad & ms
   If IsNull(rs!fIngreso) Then cad = Space(10) Else cad = Format(rs!fIngreso, "dd/mm/yyyy")
   mcad = mcad & cad & ms
   If IsNull(rs!fcese) Then cad = "11" Else cad = "13"
   mcad = mcad & cad & ms
   If rs!TipoTrabajador = "05" Then
      cad = "27"
   ElseIf rs!TipoTrabajador = "02" Then
      cad = "20"
   ElseIf rs!TipoTrabajador = "01" Then
      cad = "21"
   Else
      cad = "  "
   End If
   mcad = mcad & cad & ms
   If IsNull(rs!fcese) Then cad = Space(10) Else cad = Format(rs!fcese, "dd/mm/yyyy")
   mcad = mcad & cad & ms
   cad = lentexto(11, Left(rs!ruc_intermedio & "", 11))
   mcad = mcad & cad & ms
   If rs!essaludvida = "S" Then cad = "1" Else cad = "0"
   mcad = mcad & cad & ms

   If IsNull(rs!CodAfp) Then
      cad = "3"
   Else
      'If RS!codafp <> "" Then cad = "1" Else cad = "2"
      If rs!CodAfp = "10" Or rs!CodAfp = "11" Or rs!CodAfp = "12" Or rs!CodAfp = "17" Or rs!CodAfp = "16" Then
        cad = 1
      ElseIf rs!CodAfp = "01" Or rs!CodAfp = "02" Then
        cad = 2
        Else
            cad = 3
      End If
   End If
   mcad = mcad & cad & ms
   If IsNull(rs!sctr) Then
      cad = "0"
   Else
      If rs!sctr <> "" Then cad = "1" Else cad = "0"
   End If
   mcad = mcad & cad & ms
   If IsNull(rs!afpfechaafil) Then cad = Space(10) Else cad = Format(rs!afpfechaafil, "dd/mm/yyyy")
   mcad = mcad & cad & ms
   cad = lentexto(20, Left(rs!nomvia, 20))
   mcad = mcad & cad & ms
   cad = lentexto(4, Left(rs!nrokmmza, 4))
   mcad = mcad & cad & ms
   cad = lentexto(4, Left(rs!intdptolote, 4))
   mcad = mcad & cad & ms
   cad = lentexto(20, Left(rs!nomzona, 20))
   mcad = mcad & cad & ms
   cad = Space(40)
   mcad = mcad & cad & ms
   Select Case rs!tvia
          Case Is = "01": cad = "03"
          Case Is = "02": cad = "04"
          Case Is = "03": cad = "01"
          Case Is = "04": cad = "02"
          Case Is = "05": cad = "05"
          Case Is = "06": cad = "07"
          Case Is = "07": cad = "08"
          Case Is = "08": cad = "09"
          Case Is = "09": cad = "10"
          Case Is = "10": cad = "11"
          Case Else: cad = "12"
   End Select
   mcad = mcad & cad & ms
   
   Select Case rs!tzona
          Case Is = "01": cad = "01"
          Case Is = "02": cad = "02"
          Case Is = "03": cad = "03"
          Case Is = "04": cad = "04"
          Case Is = "05": cad = "05"
          Case Is = "06": cad = "06"
          Case Is = "07": cad = "07"
          Case Is = "08": cad = "08"
          Case Is = "09": cad = "09"
          Case Is = "10": cad = "10"
          Case Is = "11": cad = "11"
          Case Else: cad = "12"
   End Select
   mcad = mcad & cad & ms
   
   cad = lentexto(6, Mid(rs!ubigeo, 4, 6))
   mcad = mcad & cad & ms
   
   Print #1, mcad
   ProgressBar1.Value = rs.AbsolutePosition
   rs.MoveNext
Loop
Close #1
MsgBox "Se Genero el Archivo " & wruc & ".ase"
End Sub

Private Sub SpinButton1_SpinDown()
If Val(Txtano.Text) > 1 Then Txtano.Text = Txtano.Text - 1
End Sub

Private Sub SpinButton1_SpinUp()
If Val(Txtano.Text) = 0 Then Txtano.Text = "1" Else Txtano.Text = Txtano.Text + 1
End Sub

Private Sub SSCommand1_Click()
mnomcia = ""
Sql$ = "select razsoc from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then mnomcia = rs!razsoc
If rs.State = 1 Then rs.Close
If Me.Caption = "Exportando Trabajadores" Then Exporta_Trabajadores
If Me.Caption = "Exportando Remuneraciones" Then Exporta_Remuneraciones: Reporte_Remuneraciones
End Sub
Private Sub Exporta_Remuneraciones()
Dim rs2 As ADODB.Recordset
Dim ms As String
Dim td As String
Dim cad As String
Dim mcad As String
Dim mcadies As String
Dim mcadsnp As String
Dim mcadesalud As String
Dim mcadquinta As String
Dim mcadquery As String
Dim mhoras As Integer
Dim wciamae As String
ms = "|"
Dim bol_quinta As Boolean, bol_snp As Boolean

wciamae = Determina_Maestro("01076")
Sql$ = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
Sql$ = Sql$ & wciamae
mhoras = 0

If (fAbrRst(rs, Sql$)) Then mhoras = Val(rs!flag2)
If mhoras = 0 Then MsgBox "Falta Setear El numero de Horas por Periodo", vbCritical, "Importando Remuneraciones": Exit Sub
'Afecto al IES
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='02' and status<>'*'"

If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadies = ""
   Do While Not rs.EOF
      mcadies = mcadies & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   If Len(Trim(mcadies)) = 0 Then
    mcadies = "SUM(0)"
   Else
    mcadies = "SUM(" & Mid(mcadies, 1, Len(Trim(mcadies)) - 1) & ")"
   End If
Else
    mcadies = "SUM(0)"
End If

'Afecto al SNP
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='04' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadsnp = ""
   Do While Not rs.EOF
      mcadsnp = mcadsnp & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadsnp = "SUM(" & Mid(mcadsnp, 1, Len(Trim(mcadsnp)) - 1) & ")"
End If

'Afecto a Essalud
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='01' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadesalud = ""
   Do While Not rs.EOF
      mcadesalud = mcadesalud & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadesalud = "SUM(" & Mid(mcadesalud, 1, Len(Trim(mcadesalud)) - 1) & ")"
End If

'Afecto a Quinta Categoria
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='13' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadquinta = ""
   Do While Not rs.EOF
      mcadquinta = mcadquinta & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadquinta = "SUM(" & Mid(mcadquinta, 1, Len(Trim(mcadquinta)) - 1) & ")"
End If

RUTA$ = App.Path & "\REPORTS\0600" & Txtano.Text & Format(Cmbmes.ListIndex + 1, "00") & wruc & ".djt"
Open RUTA$ For Output As #1
Sql$ = "select distinct(placod),cia from plahistorico where cia='" & wcia & "' and " _
     & "year(fechaproceso)=" & Txtano.Text & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Se encuentran datos para exportar ": Exit Sub
rs.MoveFirst
ProgressBar1.Min = 0
ProgressBar1.Max = rs.RecordCount
Do While Not rs.EOF

   Sql$ = "select tipo_doc,nro_doc,quinta,codafp from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'"
   If Not (fAbrRst(rs2, Sql$)) Then
      MsgBox "No se encuentra al Codigo " & rs!PlaCod & "En la Relacion de Trabajadores " & Chr(13) & "Archivo Generado Incompleto", vbCritical, "Exportando Remuneraciones"
      If rs.State = 1 Then rs.Close
      If rs2.State = 1 Then rs2.Close
      Exit Sub
   End If
   td = "": cad = "": mcad = ""

    bol_quinta = IIf(rs2!quinta = "N", False, True)
    bol_snp = IIf(rs2!CodAfp = "01" Or rs2!CodAfp = "02" Or Len(Trim(rs2!CodAfp)) = 0, True, False)
    
   td = Trim(rs2!tipo_doc): cad = Trim(rs2!nro_doc)
   cad = lentexto(15, Left(cad, 15))
   mcad = mcad & td & ms & cad & ms
   mcadquery = ""
   mcadquery = "select " & mcadies & " as afies," & mcadsnp & " as afsnp," & mcadesalud & " as afesalud," & mcadquinta & " as afquinta,sum(d13)as quinta,sum(h01+h02+h03) as horas from plahistorico "
   mcadquery = mcadquery & "where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and "
   mcadquery = mcadquery & "year(fechaproceso)=" & Txtano.Text & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   
   If (fAbrRst(rs2, mcadquery)) Then
        Dim fecha As Date
        fecha = DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text))
      If IsNull(rs2!horas) Then cad = "  " Else cad = fCadNum(rs2!horas / mhoras, "##")
      If cad > Day(fecha) Then cad = fCadNum(Day(fecha), "##")
   End If
   mcad = mcad & cad & ms
   
   If IsNull(rs2!afies) Or Val(rs2!afies) = 0 Then cad = Space(15) Else cad = fCadNum(rs2!afies, "############.00")
   mcad = mcad & cad & ms
   
   If bol_snp Then
        If IsNull(rs2!afsnp) Then cad = Space(15) Else cad = fCadNum(rs2!afsnp, "############.00")
        mcad = mcad & cad & ms
   Else
        cad = Space(15)
        mcad = mcad & cad & ms
   End If
   
   If IsNull(rs2!afesalud) Then cad = Space(15) Else cad = fCadNum(rs2!afesalud, "############.00")
   mcad = mcad & cad & ms
   cad = Space(15)
   mcad = mcad & cad & ms
   
   If Not bol_quinta Then
       If IIf(IsNull(rs2!quinta), 0, rs2!quinta) > 0 Then
            If IsNull(rs2!afquinta) Then cad = Space(15) Else cad = fCadNum(rs2!afquinta, "############.00")
            mcad = mcad & cad & ms
            If IsNull(rs2!quinta) Then cad = Space(15) Else cad = fCadNum(rs2!quinta, "############.00")
            mcad = mcad & cad & ms
        Else
            cad = Space(15)
            mcad = mcad & cad & ms
            cad = Space(15)
            mcad = mcad & cad & ms
        End If
   Else
    If IIf(IsNull(rs2!quinta), 0, rs2!quinta) > 0 Then
        If IsNull(rs2!afquinta) Then cad = Space(15) Else cad = fCadNum(rs2!afquinta, "############.00")
        mcad = mcad & cad & ms
        If IsNull(rs2!quinta) Then cad = Space(15) Else cad = fCadNum(rs2!quinta, "############.00")
        mcad = mcad & cad & ms
    Else
        cad = Space(15)
        mcad = mcad & cad & ms
        cad = Space(15)
        mcad = mcad & cad & ms
    End If
   End If
   Print #1, mcad
   ProgressBar1.Value = rs.AbsolutePosition
   rs.MoveNext
Loop
Close #1
MsgBox "Se Genero el Archivo " & "0600" & Txtano.Text & Format(Cmbmes.ListIndex + 1, "00") & wruc & ".djt"

End Sub
Private Sub Reporte_Remuneraciones()
Dim rs2 As ADODB.Recordset
Dim ms As String
Dim td As String
Dim cad As String
Dim mcad As String
Dim mcadies As String
Dim mcadsnp As String
Dim mcadesalud As String
Dim mcadquinta As String
Dim mcadquery As String
Dim mhoras As Integer
Dim wciamae As String
Dim ties As Currency
Dim tonp As Currency
Dim tessalud As Currency
Dim tquinta As Currency
Dim tdquinta As Currency

ms = Space(1)
wciamae = Determina_Maestro("01076")
Sql$ = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
Sql$ = Sql$ & wciamae
mhoras = 0
If (fAbrRst(rs, Sql$)) Then mhoras = Val(rs!flag2)
If mhoras = 0 Then MsgBox "Falta Setear El numero de Horas por Periodo", vbCritical, "Importando Remuneraciones": Exit Sub
'Afecto al IES
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='02' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadies = ""
   Do While Not rs.EOF
      mcadies = mcadies & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadies = "SUM(" & Mid(mcadies, 1, Len(Trim(mcadies)) - 1) & ")"
Else
    mcadies = "SUM(0)"
End If

'Afecto al SNP
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='04' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadsnp = ""
   Do While Not rs.EOF
      mcadsnp = mcadsnp & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadsnp = "SUM(" & Mid(mcadsnp, 1, Len(Trim(mcadsnp)) - 1) & ")"
End If

'Afecto a Essalud
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='01' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadesalud = ""
   Do While Not rs.EOF
      mcadesalud = mcadesalud & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadesalud = "SUM(" & Mid(mcadesalud, 1, Len(Trim(mcadesalud)) - 1) & ")"
End If

'Afecto a Quinta Categoria
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='13' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadquinta = ""
   Do While Not rs.EOF
      mcadquinta = mcadquinta & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadquinta = "SUM(" & Mid(mcadquinta, 1, Len(Trim(mcadquinta)) - 1) & ")"
End If

RUTA$ = App.Path & "\REPORTS\ReportRemu.txt"
Open RUTA$ For Output As #1
Sql$ = "select distinct(placod),cia from plahistorico where cia='" & wcia & "' and " _
     & "year(fechaproceso)=" & Txtano.Text & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Se encuentran datos para exportar ": Exit Sub
rs.MoveFirst
ProgressBar1.Min = 0
ProgressBar1.Max = rs.RecordCount
mpag = 0
ties = 0: tonp = 0: tessalud = 0: tquinta = 0: tdquinta = 0
Cabeza_ReporteRemu

Do While Not rs.EOF
   Sql$ = "select ap_pat,ap_mat,nom_1,nom_2,dni,lmilitar,pasaporte,tipo_doc,nro_doc from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'"
   If Not (fAbrRst(rs2, Sql$)) Then
      MsgBox "No se encuentra al Codigo " & rs!PlaCod & "En la Relacion de Trabajadores " & Chr(13) & "Archivo Generado Incompleto", vbCritical, "Exportando Remuneraciones"
      If rs.State = 1 Then rs.Close
      If rs2.State = 1 Then rs2.Close
      Exit Sub
   End If
   td = "": cad = "": mcad = ""
   mcad = rs!PlaCod & ms
   cad = lentexto(15, Left(rs2!ap_pat, 15))
   mcad = mcad & ms & cad & ms
   cad = lentexto(15, Left(rs2!ap_mat, 15))
   mcad = mcad & ms & cad & ms
   cad = lentexto(15, Left(rs2!nom_1, 15))
   mcad = mcad & ms & cad & ms
   
   td = rs2!tipo_doc: cad = rs2!nro_doc
   
   cad = lentexto(15, Left(cad, 15))
   mcad = mcad & td & ms & cad & ms
   mcadquery = ""
   mcadquery = "select " & mcadies & " as afies," & mcadsnp & " as afsnp," & mcadesalud & " as afesalud," & mcadquinta & " as afquinta,sum(d13)as quinta,sum(h01+h02+h03) as horas from plahistorico "
   mcadquery = mcadquery & "where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and "
   mcadquery = mcadquery & "year(fechaproceso)=" & Txtano.Text & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   
   If (fAbrRst(rs2, mcadquery)) Then
        Dim fecha As Date
        fecha = DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text))
      If IsNull(rs2!horas) Then cad = "  " Else cad = fCadNum(rs2!horas / mhoras, "##")
      If cad > Day(fecha) Then cad = fCadNum(Day(fecha), "##")
   End If
   
   mcad = mcad & cad & ms
   If IsNull(rs2!afies) Then cad = Space(15) Else cad = fCadNum(rs2!afies, "###,###.00")
   mcad = mcad & cad & ms
   
   If IsNull(rs2!afsnp) Then cad = Space(15) Else cad = fCadNum(rs2!afsnp, "###,###.00")
   mcad = mcad & cad & ms
   
   If IsNull(rs2!afesalud) Then cad = Space(15) Else cad = fCadNum(rs2!afesalud, "###,###.00")
   mcad = mcad & cad & ms
   If IsNull(rs2!afquinta) Then cad = Space(15) Else cad = fCadNum(rs2!afquinta, "###,###.00")
   mcad = mcad & cad & ms
   If IsNull(rs2!quinta) Then cad = Space(15) Else cad = fCadNum(rs2!quinta, "###,###.00")
   mcad = mcad & cad
   Print #1, Space(2) & mcad
   ties = ties + rs2!afies
   tonp = tonp + rs2!afsnp
   tessalud = tessalud + rs2!afesalud
   tquinta = tquinta + rs2!afquinta
   tdquinta = tdquinta + rs2!quinta
   mlinea = mlinea + 1
   If mlinea > 55 Then Print #1, SaltaPag: Cabeza_ReporteRemu
   ProgressBar1.Value = rs.AbsolutePosition
   rs.MoveNext
Loop
Print #1, Space(2) & String(134, "-")
Print #1, Space(81) & fCadNum(ties, "###,###.00") & ms & fCadNum(tonp, "###,###.00") & ms & fCadNum(tessalud, "###,###.00") & ms & fCadNum(tquinta, "###,###.00") & ms & fCadNum(tdquinta, "###,###.00") & ms
Close #1
Call Imprime_Txt("ReportRemu.txt", RUTA$)
End Sub
Private Sub Cabeza_ReporteRemu()
mpag = mpag + 1
Print #1, LetraChica & Space(2) & mnomcia
Print #1, Space(32) & "Datos a Transferir a la SUNAT" & Space(65) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(2) & String(134, "-")
Print #1, "  CODIGO             APELLIDOS Y NOMBRES                  DOC. IDENTIDAD      DT   REM. IES   REM. ONP  REM.ESSALUD REM. QTA      QTA."
Print #1, Space(2) & String(134, "-")
mlinea = 7
End Sub
