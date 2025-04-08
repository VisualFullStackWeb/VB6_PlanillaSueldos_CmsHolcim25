VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmsubsidios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subsidios y Liquidaciones"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "Frmsubsidios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      ItemData        =   "Frmsubsidios.frx":030A
      Left            =   1320
      List            =   "Frmsubsidios.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   2400
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Txtcese 
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker Cmbfecha 
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   60358657
      CurrentDate     =   37662
   End
   Begin VB.TextBox Txtimporte 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Txtcodpla 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   2895
      Begin VB.OptionButton OpcQta 
         Caption         =   "Quinta"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   420
         Width           =   1215
      End
      Begin VB.OptionButton Opcliquid 
         Caption         =   "Liquidacion"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Opcsubs 
         Caption         =   "Subsidio x enfermedad"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   2535
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
      Width           =   5295
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   3975
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Boleta"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label Lblfnac 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblfnac"
      Height          =   195
      Left            =   5640
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Lblarea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblarea"
      Height          =   195
      Left            =   5640
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Lblbasico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblbasico"
      Height          =   195
      Left            =   5640
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Lblnumafp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblnumafp"
      Height          =   195
      Left            =   5640
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lblobra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblobra"
      Height          =   195
      Left            =   5640
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Lbltipotrab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lbltipotrab"
      Height          =   195
      Left            =   5640
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lblafp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblafp"
      Height          =   195
      Left            =   5640
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Lblfingreso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblfingreso"
      Height          =   195
      Left            =   5640
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Lblcodaux 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lblcodaux"
      Height          =   195
      Left            =   5640
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lblcese 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fec. Cese"
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   720
   End
   Begin MSForms.CheckBox Chkquinta 
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1508;450"
      Value           =   "0"
      Caption         =   "Quinta"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      TabIndex        =   12
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Lblnombre 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   765
   End
End
Attribute VB_Name = "Frmsubsidios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim manos As Integer
Dim VTipobol As String
Dim VSemana As String
Dim Vano As Integer
Dim Vmes As Integer
Dim macui As Currency
Dim macus As Currency
Dim VFProceso As String

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5415
Me.Height = 3495
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Cmbfecha.Value = Date
Limpiar
Txtcodpla.Text = ""

If wGrupoPla = "01" And UCase(wuser) <> "SA" Then
   wciamae = Determina_Maestro("01078")
   Sql$ = "Select COD_MAESTRO2,DESCRIP from maestros_2 where status<>'*' and (cod_maestro2 in(select tipo from pla_permisos where usuario='" & wuser & "' and calculo='B'))"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then rs.MoveFirst
   Do While Not rs.EOF
      Cmbtipo.AddItem rs!DESCRIP
      Cmbtipo.ItemData(Cmbtipo.NewIndex) = Trim(rs!cod_maestro2)
      rs.MoveNext
   Loop
   rs.Close
Else
   Call fc_Descrip_Maestros2("01078", "", Cmbtipo)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Frmgrdsubsidio.Procesa_Cabeza_Subsidio
End Sub

Private Sub Opcliquid_Click()
If Opcliquid.Value = True Then
   LblCese.Visible = True
   Txtcese.Visible = True
   Chkquinta.Visible = True
End If
Limpiar
End Sub

Private Sub OpcQta_Click()
If Opcsubs.Value = True Then
   LblCese.Visible = False
   Txtcese.Visible = False
   Chkquinta.Visible = False
End If
Limpiar
End Sub

Private Sub Opcsubs_Click()
If Opcsubs.Value = True Then
   LblCese.Visible = False
   Txtcese.Visible = False
   Chkquinta.Visible = False
End If
Limpiar
End Sub

Public Sub Txtcodpla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtImporte.SetFocus
End Sub
Private Sub Txtcodpla_LostFocus()
Dim mcad As String
Limpiar
If Trim(Txtcodpla.Text) = "" Then Exit Sub
cod = "01055"
Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(rs, Sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If rs.State = 1 Then rs.Close

Sql$ = nombre()
Sql$ = Sql$ + "codauxinterno,a.status,a.tipotrabajador,a.obra,a.fingreso,a.fcese,a.codafp,a.numafp,a.area,a.placod,a.codauxinterno,b.descrip,a.tipotasaextra,a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento,a.fec_jubila " _
     & "from planillas a,maestros_2 b where a.status<>'*'"
     Sql$ = Sql$ & xciamae
     Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 " _
     & "and cia='" & wcia & "' AND placod='" & Txtcodpla.Text & "'" & xciamae
     Sql$ = Sql$ & " order by nombre"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$)
If rs.RecordCount > 0 Then
   If Opcsubs.Value = True Or OpcQta.Value = True Then
      If Not IsNull(rs!fcese) Then
         MsgBox "Trabajador ya fue Cesado", vbExclamation, "Con Fecha => " & Format(rs!fcese, "dd/mm/yyyy")
         Txtcodpla.Text = ""
         Exit Sub
      End If
   Else
      If IsNull(rs!fcese) Then
         Txtcese.Enabled = True
      Else
         Txtcese.Text = Format(rs!fcese, "dd/mm/yyyy")
         Txtcese.Enabled = False
      End If
   End If
   Lblnombre.Caption = rs!nombre
   Lblcodaux.Caption = rs!codauxinterno
   LblFingreso.Caption = Format(rs!fIngreso, "mm/dd/yyyy")
   Lblfnac.Caption = Format(rs!fnacimiento, "mm/dd/yyyy")
   Lblafp.Caption = rs!CodAfp
   LblTipoTrab.Caption = rs!TipoTrabajador
   Lblobra.Caption = rs!obra
   Lblnumafp.Caption = rs!NUMAFP
   Lblarea.Caption = rs!Area
Else
   If Opcsubs.Value = True Then mcad = "Subsidio" Else If OpcQta.Value = True Then mcad = "Quinta" Else mcad = "Liquidacion"
   MsgBox "Codigo de Trabajador no Registrado", vbInformation, mcad
   mcad = ""
   Limpiar
   Txtcodpla.Text = ""
End If
If rs.State = 1 Then rs.Close
'Basico
Sql$ = "select importe from plaremunbase where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and concepto='01' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then LblBasico.Caption = rs!importe
If rs.State = 1 Then rs.Close
End Sub

Private Sub Txtimporte_KeyPress(KeyAscii As Integer)
TxtImporte.Text = TxtImporte.Text + Fc_Decimals(KeyAscii)
End Sub

Public Sub Grabar_Subsidio()
Dim NroTrans As Integer
Dim MqueryCalD As String
Dim MqueryCalA As String
Dim mtoting As Currency
Dim itemcosto As Integer
Dim mcad As String
On Error GoTo ErrorTrans
NroTrans = 0

mcancel = False
mtoting = 0
VFProceso = Cmbfecha.Value
If Opcsubs.Value = True Then mcad = "Subsidio" Else If OpcQta.Value = True Then mcad = "Quinta" Else mcad = "Liquidacion"
If Not IsNumeric(TxtImporte.Text) Then MsgBox "Debe Indicar un Monto Correcto", vbCritical, TitMsg: Cmbturno.SetFocus: Exit Sub
Mgrab = MsgBox("Seguro de Grabar " & mcad, vbYesNo + vbQuestion, mcad)
If Mgrab <> 6 Then Exit Sub
mcad = ""
Screen.MousePointer = vbArrowHourglass
Vano = Cmbfecha.Year
Vmes = Cmbfecha.Month
manos = perendat(Cmbfecha.Value, Lblfnac.Caption, "a")
If Opcsubs.Value = True Then VTipobol = "04" Else VTipobol = "05"

cn.BeginTrans
NroTrans = 1
Sql$ = "insert into platemphist(cia,placod,codauxinterno,proceso,fechaproceso,semana,fechaingreso,turno,codafp,status,fec_crea,tipotrab,obra,numafp,basico,ccosto1,porc1,i01,fec_modi,user_crea,user_modi) " _
     & "values('" & wcia & "','" & Txtcodpla.Text & "','" & Lblcodaux.Caption & "','" & VTipobol & "','" & Format(Cmbfecha.Value, FormatFecha) & "','" & VSemana & "','" & Format(LblFingreso.Caption, FormatFecha) & "','','" & Lblafp.Caption & "','T'," & FechaSys & ",'" & LblTipoTrab.Caption & "','" & Lblobra.Caption & "','" & Lblnumafp.Caption & "'," & CCur(LblBasico.Caption) & ",'" & Lblarea.Caption & "',100.00," & CCur(TxtImporte.Text) & "," & FechaSys & ",'" & wuser & "','" & wuser & "')"
cn.Execute Sql

mcad = ""

'Calculo de Deducciones
MqueryCalD = ""
For I = 1 To 0
    Sql$ = F02(Format(I, "00"))
    If mcancel = True Then
       MsgBox "Se Cancelo la Grabacion", vbCritical, "Calculo de Boleta"
       Sql$ = wCancelTrans
       cn.Execute Sql$
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If Sql$ <> "" Then
       If (fAbrRst(rs, Sql$)) Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
          Else
             If I = 11 Then
                For J = 1 To 5
                   MqueryCalD = MqueryCalD & "d" & Format(I, "00") & Format(J, "0") & " = " & rs(J - 1) & ","
                Next J
             Else
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
             End If
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next I
If MqueryCalD <> "" Then
   MqueryCalD = Mid(MqueryCalD, 1, Len(Trim(MqueryCalD)) - 1)
   Sql$ = "Update platemphist set " & MqueryCalD
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='T'"
   cn.Execute Sql
End If

'Calculo de Aportaciones
MqueryCalA = ""
For I = 1 To 20
    Sql$ = F03(Format(I, "00"))
    If Sql$ <> "" Then
       cn.CursorLocation = adUseClient
       Set rs = New ADODB.Recordset
       Set rs = cn.Execute(Sql$, 64)
       If rs.RecordCount > 0 Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
          Else
             MqueryCalA = MqueryCalA & "a" & Format(I, "00") & " = " & rs(0) & ","
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next
If MqueryCalA <> "" Then
   MqueryCalA = Mid(MqueryCalA, 1, Len(Trim(MqueryCalA)) - 1)
   Sql$ = "Update platemphist set " & MqueryCalA
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='T'"
   cn.Execute Sql
End If

Dim mi As String, md As String, ma As String, mi2 As String
mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
'md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
'ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
'add jcms 260921 corrige add new campo de calculo
md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20+d21"
ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20+a21"

mi2 = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+ " & TxtImporte.Text & "+i44+i45+i46+i47+i48+i49+i50"
md2 = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+ " & TxtImporte.Text & " +d14+d15+d16+d17+d18+d19+d20+d21"

Sql$ = "update platemphist set totaling=" & mi & "," _
  & "d11=d111+d112+d113+d114+d115," _
  & "totalded=" & md & "," _
  & "totalapo=" & ma & "," _
  & "totneto=(" & mi & ")-" & "(" & md & ")"
Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='T'"

cn.Execute Sql

'Sql$ = "insert into plahistorico select * from platemphist"
'Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
'Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='T'"
'cn.Execute Sql
'Cambio Mgirao para que no genere una boleta y modifique la del plahistorico del mes con el importe del subsidio
Dim VTipo As String
VTipo = fc_CodigoComboBox(Cmbtipo, 2)

If Opcsubs.Value = True Then
    Sql$ = "update plahistorico set i43= " & TxtImporte.Text & " ,totaling=" & mi2 & "," _
    & "totalded=" & md & "," _
        & "totalapo=" & ma & "," _
        & "totneto=(" & mi2 & ")-" & "(" & md & ")"
    Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipo & "' "
    'Sql$ = Sql$ & "and year(fechaproceso)='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status<>'*'"
    Sql$ = Sql$ & " and year(fechaproceso)=" & Vano & " and month(fechaproceso)='" & Vmes & "' and semana='" & VSemana & "' and status<>'*'"
    
End If
If OpcQta.Value = True Then
    Sql$ = "update plahistorico set d13= " & TxtImporte.Text & " ,totalded=" & md2 & "," _
    & "totaling=" & mi & "," _
        & "totalapo=" & ma & "," _
        & "totneto=(" & mi & ")-" & "(" & md2 & ")"
    Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='01' "
    'Sql$ = Sql$ & "and year(fechaproceso)='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status<>'*'"
    Sql$ = Sql$ & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)='" & Vmes & "' and semana='" & VSemana & "' and status<>'*'"
End If
cn.Execute Sql


Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
Sql$ = Sql$ & " and cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='T'"

cn.Execute Sql

If Opcliquid.Value = True And Txtcese.Enabled = True Then
   Sql$ = "Update planillas set fcese='" & Format(Txtcese, FormatFecha) & "' where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and status<>'*'"
   cn.Execute Sql$
End If

cn.CommitTrans
Limpiar
Txtcodpla.Text = ""
MsgBox "Se Grabó el Subsidio Satisfactoriamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

Limpiar
Txtcodpla.Text = ""
MsgBox Err.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault


End Sub
Public Sub Limpiar()
'Txtcodpla.Text = ""
Lblnombre.Caption = ""
Lblcodaux.Caption = ""
LblFingreso.Caption = ""
Lblafp.Caption = ""
LblTipoTrab.Caption = ""
Lblobra.Caption = ""
Lblnumafp.Caption = ""
LblBasico.Caption = ""
Lblarea.Caption = ""
TxtImporte.Text = ""
Txtcese.Text = "__/__/____"
Chkquinta.Value = 0
Lblfnac.Caption = ""
End Sub
Public Function F02(concepto As String) As String 'DEDUCCIONES
Dim rsF02 As ADODB.Recordset
Dim rsF02afp As ADODB.Recordset
Dim F02str As String
Dim rsTope As ADODB.Recordset
Dim mFactor As Currency
Dim mperiodoafp As String
Dim vNombField As String
Dim mtope As Currency
Dim mproy As Currency
Dim MUIT As Currency
Dim mgra As Integer
Dim msemano As Integer
Dim mpertope As Integer
Dim J As Integer
mFactor = 0
F02 = ""
mtope = 0
If (concepto <> "04" Or Lblafp.Caption = "") And concepto <> "11" And concepto <> "13" Then 'SIN AFP
   If Not IsDate(VFechaJub) Then
    Sql$ = "select deduccion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and deduccion<>0 and status<>'*'"
    If (fAbrRst(rsF02, Sql$)) Then
       If Not IsNull(rsF02!deduccion) Then
          If rsF02!deduccion <> 0 Then mFactor = rsF02!deduccion
       End If
    End If
    If rsF02.State = 1 Then rsF02.Close
    If mFactor <> 0 Then
       Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
       If (fAbrRst(rsF02, Sql$)) Then
          rsF02.MoveFirst
          F02str = ""
          Do While Not rsF02.EOF
             F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
             rsF02.MoveNext
          Loop
          F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
          If rsF02.State = 1 Then rsF02.Close
          Call Acumula_Mes(concepto, "D")
          F02 = "select round(((" & F02str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as deduccion from platemphist "
          F02 = F02 & "where cia='" & wcia & "' and status='T' and placod='" & Txtcodpla.Text & "' "
          F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
       End If
    End If
   End If
ElseIf concepto = "11" And Lblafp.Caption <> "" Then 'AFP
   If Not IsDate(VFechaJub) Then
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(rsF02, Sql$)) Then
       rsF02.MoveFirst
       F02str = ""
       Do While Not rsF02.EOF
          F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
          rsF02.MoveNext
       Loop
       F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
       
       If rsF02.State = 1 Then rsF02.Close
       mperiodoafp = Format(Cmbfecha.Year, "0000") & Format(Cmbfecha.Month, "00")
       Sql$ = "select afp01,afp02,afp03,afp04,afp05,tope from  plaafp where periodo='" & mperiodoafp & "' and codafp='" & Lblafp.Caption & "' and status<>'*'"
       If Not (fAbrRst(rsF02, Sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para AFP", vbCritical, "Calculo de Boleta"
          mcancel = True
          Exit Function
       End If
       Sql$ = Acumula_Mes_Afp(concepto, "D")
       If (fAbrRst(rsF02afp, Sql$)) Then
          For J = 1 To 5
              vNombField = " as D11" & Format(J, "0")
              mFactor = rsF02(J - 1)
              If J = 2 Then
                 If manos > 64 Then mFactor = 0
                 Call Acumula_Mes_Afp112(concepto, "D")
                 mtope = macui
                 Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='T' and placod='" & Txtcodpla.Text & "' " _
                      & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
                      
                 If (fAbrRst(rsTope, Sql$)) Then mtope = mtope + rsTope!tope
                 If mtope > rsF02!tope Then mtope = rsF02!tope
                 If rsTope.State = 1 Then rsTope.Close
                 F02 = F02 & "round(((" & mtope & ") * " & mFactor & " /100)-" & macus & ",2) "
                 F02 = F02 & vNombField & ","
              Else
                 If Not IsNull(rsF02afp(0)) Then
                    F02 = F02 & "round(((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & ",2) "
                    F02 = F02 & vNombField & ","
                 Else
                    F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                    F02 = F02 & vNombField & ","
                 End If
              End If
          Next J
          If rsF02afp.State = 1 Then rsF02afp.Close
          If rsF02.State = 1 Then rsF02.Close
       End If
       F02 = Mid(F02, 1, Len(Trim(F02)) - 1)
       F02 = "select " & F02
       F02 = F02 & " from platemphist "
       F02 = F02 & "where cia='" & wcia & "' and status='T' and placod='" & Txtcodpla.Text & "' "
       F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
    End If
   End If
ElseIf concepto = "13" Then 'Quinta Categoria
   If Opcsubs.Value = True Or Chkquinta.Value = 1 Or OpcQta.Value = True Then
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF02, Sql$)) Then
      rsF02.MoveFirst
      F02str = ""
      Do While Not rsF02.EOF
         F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
         rsF02.MoveNext
      Loop
      F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      MUIT = 0
      mtope = 0
      mpertope = 0
      mgra = 0
      msemano = 0
      If vTipoTra <> "01" Then
         Sql$ = "select max(semana) as ultima from plasemanas where cia='" & wcia & "' and ano='" & Format(Vano, "0000") & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mpertope = rsF02(0): msemano = rsF02(0)
      Else
         If Vmes > 6 Then mpertope = 12 Else mpertope = 13
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      Call Acumula_Ano(concepto, "D")
      mtope = macui
      Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='T' and placod='" & Txtcodpla.Text & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      If (fAbrRst(rsF02, Sql$)) Then mtope = mtope + rsF02!tope
      If rsF02.State = 1 Then rsF02.Close
      
'      Sql$ = "select concepto,moneda,sum((importe/factor_horas)) as base from plaremunbase a,plaafectos b where a.cia='" & wcia & "' " _
'           & "and placod='" & Txtcodpla.Text & "' and a.status<>'*' and b.tipo='D' and b.codigo='" & concepto & "' and b.tboleta='01' and b.status<>'*' and a.concepto=b.cod_remu " _
'           & "Group By Placod"

      ' MOD LFSA - 14/02/2013
      Sql$ = "select sum((importe/factor_horas)) as base from plaremunbase a,plaafectos b where a.cia='" & wcia & "' " _
           & "and placod='" & Txtcodpla.Text & "' and a.status<>'*' and b.tipo='D' and b.codigo='" & concepto & "' and b.tboleta='01' and b.status<>'*' and a.concepto=b.cod_remu " _
           & "Group By Placod"
           
      If (fAbrRst(rsF02, Sql$)) Then mproy = rsF02!base
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select uit from plauit where ano='" & Format(Vano, "0000") & "' and moneda='S/.' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then MUIT = rsF02!uit
      If rsF02.State = 1 Then rsF02.Close
      
      If vTipoTra = "05" Then
         Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='20' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mFactor = rsF02!factor
         If rsF02.State = 1 Then rsF02.Close
      
         Sql$ = "select concepto,moneda,importe/factor_horas as base from plaremunbase where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and concepto='01' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mproy = mproy + Round(rsF02!base * mFactor, 2)
         If rsF02.State = 1 Then rsF02.Close
      End If
      If vTipoTra = "01" Then
         If Vmes < 12 Then mproy = (mproy * VHoras) * (mpertope - Vmes + 1) Else mproy = 0
      Else
         mgra = Busca_Grati()
         If vTipoTra = "05" Then
            Sql$ = "select importe/factor_horas as base,b.factor  from plaremunbase a,platasaanexo b where a.cia='" & wcia & "' and a.placod='" & Txtcodpla.Text & "'  and a.concepto='01' " _
                 & "and a.status<>'*' and b.cia='" & wcia & "' and b.tipomovimiento='01' and b.codinterno='15' and b.status<>'*' and b.tipotrab='" & vTipoTra & "' and b.cargo='" & Trim(Lblcargo.Caption) & "'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((rsF02!base * 8) * (rsF02!factor * mgra)) + ((rsF02!base * 8) * (mpertope - Val(VSemana)))
            If rsF02.State = 1 Then rsF02.Close
         Else
            Sql$ = "select importe/factor_horas as base from plaremunbase where Cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and concepto='01' and status<>'*'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (rsF02!base * 8) * (mpertope - Val(VSemana))
            If rsF02.State = 1 Then rsF02.Close
            'mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (mproy * 8) * (mpertope - Val(VSemana))
         End If
      End If
      mtope = mtope + mproy
      If mtope > Round(MUIT * 7, 2) Then
         mtope = mtope - Round(MUIT * 7, 2)
         Select Case mtope
                Case Is < (Round(MUIT * 27, 2) + 1)
                     mFactor = Round(mtope * 0.15, 2)
                Case Is < (Round(MUIT * 54, 2) + 1)
                     mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
                Case Else
                     mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
         End Select
         If vTipoTra = "01" Then
            F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
         Else
            If VTipobol = "02" Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
            Else
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (" & msemano & " - " & Val(VSemana) & " + 1), 2)"
            End If
         End If
      End If
   End If
   End If
End If
macui = 0: macus = 0
End Function
Public Function F03(concepto As String) As String 'APORTACIONES
Dim rsF03 As ADODB.Recordset
Dim F03str As String
Dim mFactor As Currency
mFactor = 0
F03 = ""
If concepto = "03" Then
   Sql$ = "select senati from cia where cod_cia='" & wcia & "' and status<>'*'"
   If Not (fAbrRst(rsF03, Sql$)) Then Exit Function
   If rsF03!senati <> "S" Then Exit Function
   If rsF03.State = 1 Then rsF03.Close
   wciamae = Determina_Maestro("01044")
   Sql$ = "Select * from maestros_2 where cod_maestro2='" & VArea & "' and status<>'*'"
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rsF03, Sql$)) Then
      If rsF03!flag7 <> "S" Then Exit Function
   Else
      Exit Function
   End If
   If rsF03.State = 1 Then rsF03.Close
End If

Sql$ = "select aportacion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and aportacion<>0 and status<>'*'"
If (fAbrRst(rsF03, Sql$)) Then
   If Not IsNull(rsF03!aportacion) Then
      If rsF03!aportacion <> 0 Then mFactor = rsF03!aportacion
   End If
End If
If rsF03.State = 1 Then rsF03.Close
If mFactor <> 0 Then
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF03, Sql$)) Then
      rsF03.MoveFirst
      F03str = ""
      Do While Not rsF03.EOF
         F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
         rsF03.MoveNext
      Loop
      F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
      If rsF03.State = 1 Then rsF03.Close
      Call Acumula_Mes(concepto, "A")
      F03 = "select round(((" & F03str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as aportacion from platemphist "
      F03 = F03 & "where cia='" & wcia & "' and status='T' and placod='" & Txtcodpla.Text & "' "
      F03 = F03 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
   End If
End If

macui = 0: macus = 0
End Function


Private Sub Acumula_Mes(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim mtb2 As String
macui = 0: macus = 0
For I = 1 To 5
    If I = 1 Then mtb = "01": mtb2 = "01"
    If I = 2 Then mtb = "02": mtb2 = "02"
    If I = 3 Then mtb = "03": mtb2 = "03"
    If I = 4 Then mtb = "04": mtb2 = "01"
    If I = 5 Then mtb = "05": mtb2 = "01"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb2 & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub

Private Sub Acumula_Ano(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim mtb2 As String
macui = 0: macus = 0
For I = 1 To 5
    If I = 1 Then mtb = "01": mtb2 = "01"
    If I = 2 Then mtb = "02": mtb2 = "02"
    If I = 3 Then mtb = "03": mtb2 = "03"
    If I = 4 Then mtb = "04": mtb2 = "01"
    If I = 5 Then mtb = "05": mtb2 = "01"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb2 & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)<=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub

Private Function Acumula_Mes_Afp(concepto As String, tipo As String) As String
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d111) as ded1, " _
              & "sum(d112) as ded2, " _
              & "sum(d113) as ded3, " _
              & "sum(d114) as ded4, " _
              & "sum(d115) as ded5 "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
    End If
If RsAcumula.State = 1 Then RsAcumula.Close
Acumula_Mes_Afp = SqlAcu
End Function


Private Sub Acumula_Mes_Afp112(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim mtb2 As String
macui = 0: macus = 0
For I = 1 To 5
    If VTipobol = "02" And I <> 2 Then I = I + 1
    If VTipobol <> "02" And I = 2 Then I = I + 1
    If I > 3 Then Exit For
    If I = 1 Then mtb = "01": mtb2 = "01"
    If I = 2 Then mtb = "02": mtb2 = "02"
    If I = 3 Then mtb = "03": mtb2 = "03"
    If I = 4 Then mtb = "03": mtb2 = "01"
    If I = 5 Then mtb = "03": mtb2 = "01"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb2 & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d112) as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub

Private Function Busca_Grati() As Integer
Dim rsGrati As New Recordset
Select Case Vmes
       Case Is = 1, 2, 3, 4, 5, 6
            Busca_Grati = 2
       Case Is = 7
            Sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 1 Else Busca_Grati = 2
                 
       Case Is = 8, 9, 10, 11
            Busca_Grati = 1
       Case Is = 12
            Sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 0 Else Busca_Grati = 1
End Select
End Function


