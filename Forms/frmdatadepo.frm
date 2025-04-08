VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{92F3EE28-0DF4-4CB0-AAA5-E6578B12929B}#1.0#0"; "CboFacil.ocx"
Begin VB.Form frmdatadepo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Archivo PAGHAB.DAT"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txt_Semana 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      ItemData        =   "frmdatadepo.frx":0000
      Left            =   120
      List            =   "frmdatadepo.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Archivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5880
      Width           =   6495
   End
   Begin CboFacil.cbo_facil cbomoneda 
      Height          =   315
      Left            =   2175
      TabIndex        =   2
      Top             =   990
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      NameTab         =   ""
      NameCod         =   ""
      NameDesc        =   ""
      Filtro          =   ""
      OrderBy         =   ""
      SetIndex        =   ""
      NameSistema     =   ""
      Mensaje         =   0   'False
      ToolTip         =   0   'False
      Enabled         =   -1  'True
      Ninguno         =   0   'False
      BackColor       =   -2147483643
      BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CboFacil.cbo_facil cbobanco 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   556
      NameTab         =   ""
      NameCod         =   ""
      NameDesc        =   ""
      Filtro          =   ""
      OrderBy         =   ""
      SetIndex        =   ""
      NameSistema     =   ""
      Mensaje         =   0   'False
      ToolTip         =   0   'False
      Enabled         =   -1  'True
      Ninguno         =   0   'False
      BackColor       =   -2147483643
      BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   37265
   End
   Begin vbAcceleratorSGrid6.vbalGrid vbgcts 
      Height          =   3750
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   6615
      NoVerticalGridLines=   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      GridLineColor   =   15466236
      HighlightBackColor=   15466236
      HighlightForeColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      HighlightSelectedIcons=   0   'False
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   5580
      Top             =   225
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   14924
      Images          =   "frmdatadepo.frx":002A
      Version         =   131072
      KeyCount        =   13
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin CboFacil.cbo_facil cbo_boletas 
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      NameTab         =   ""
      NameCod         =   ""
      NameDesc        =   ""
      Filtro          =   ""
      OrderBy         =   ""
      SetIndex        =   ""
      NameSistema     =   ""
      Mensaje         =   0   'False
      ToolTip         =   0   'False
      Enabled         =   -1  'True
      Ninguno         =   0   'False
      BackColor       =   -2147483643
      BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   1560
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sem"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
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
      Top             =   765
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Pago"
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
      Left            =   4320
      TabIndex        =   11
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
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
      Left            =   2175
      TabIndex        =   6
      Top             =   765
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
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
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   510
   End
End
Attribute VB_Name = "frmdatadepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CONS_EMPRESA As String, CONS_CUENTA_DET As String, CONS_FIJODET As String
Dim SUMIMPORTES As Currency, sumacta As Variant, ultfila As Long
Dim VTipotrab As String, wciamae As String
Const BOLETA_NORMAL = 1
Const BOLETA_VACACIONES = 2
Const BOLETA_GRATIFICACION = 3

Public Sub Procesar()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim CADENAPRINT As String, CADENAPRINTDET As String
Dim ARRDET() As Variant, IDENT As String, RUTA As String
Dim MAXROW As Long

'On Error GoTo Archivo

Sql = "SELECT * FROM constantes_bancos WHERE CIA='" & wcia & "' "
Set Rs = cn.Execute(Sql)

If Not Rs.EOF Then
    CONS_EMPRESA = Rs(1)
    CONS_CUENTA_DET = Rs(2)
    CONS_FIJODET = Rs(3)
    Rs.Close
End If

If cbo_boletas.ReturnCodigo = BOLETA_NORMAL Then
    If Day(txtfecha.Value) > 15 Then
        Sql = "SELECT a.PLACOD, a.totneto, b.sucursal, pagonumcta, pagotipcta, pagomoneda, LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2)) AS NOMBRE FROM PLAHISTORICO "
        Sql = Sql & " a INNER JOIN planillas b on (b.cia=a.cia and b.placod=a.placod and b.status!='*' and b.pagomoneda='" & Mid(cbomoneda.Text, 2, 3) & "' "
        Sql = Sql & "  and pagobanco='" & Format(cbobanco.ReturnCodigo, "00") & "') WHERE a.CIA='" & wcia & "' AND a.STATUS!='*' AND a.PROCESO='01' and "
        Sql = Sql & " month(a.fechaproceso)=" & Month(txtfecha.Value) & " and year(a.fechaproceso)=" & Year(txtfecha.Value) & " and pagonumcta!='0' "
        Sql = Sql & " AND A.TIPOTRAB='" & VTipotrab & "' AND A.SEMANA ='" & Trim(Txt_Semana.Text) & "'"
        Sql = Sql & " ORDER BY LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2))"
        
    Else
        Sql = "SELECT a.PLACOD, a.totneto, b.sucursal, pagonumcta, pagotipcta, pagomoneda, LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2)) AS NOMBRE FROM PLAQUINCENA "
        Sql = Sql & " a INNER JOIN planillas b on (b.cia=a.cia and b.placod=a.placod and b.status!='*' and b.pagomoneda='" & Mid(cbomoneda.Text, 2, 3) & "' "
        Sql = Sql & "  and pagobanco='" & Format(cbobanco.ReturnCodigo, "00") & "') WHERE a.CIA='" & wcia & "' AND a.STATUS!='*' AND a.PROCESO='01' and "
        Sql = Sql & " month(a.fechaproceso)=" & Month(txtfecha.Value) & " and year(a.fechaproceso)=" & Year(txtfecha.Value) & " and pagonumcta!='0' "
        Sql = Sql & " AND A.TIPOTRAB=' " & VTipotrab & "' AND A.SEMANA =' " & Trim(Txt_Semana.Text) & "'"
        Sql = Sql & "ORDER BY LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2))"
    End If
ElseIf cbo_boletas.ReturnCodigo = BOLETA_VACACIONES Then
        Sql = "SELECT a.PLACOD, a.totneto, b.sucursal, pagonumcta, pagotipcta, pagomoneda, LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2)) AS NOMBRE FROM PLAHISTORICO "
        Sql = Sql & " a INNER JOIN planillas b on (b.cia=a.cia and b.placod=a.placod and b.status!='*' and b.pagomoneda='" & Mid(cbomoneda.Text, 2, 3) & "' "
        Sql = Sql & "  and pagobanco='" & Format(cbobanco.ReturnCodigo, "00") & "') WHERE a.CIA='" & wcia & "' AND a.STATUS!='*' AND a.PROCESO='02' and "
        Sql = Sql & " month(a.fechaproceso)=" & Month(txtfecha.Value) & " and year(a.fechaproceso)=" & Year(txtfecha.Value) & " and pagonumcta!='0' ORDER BY LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2))"
ElseIf cbo_boletas.ReturnCodigo = BOLETA_GRATIFICACION Then
        Sql = "SELECT a.PLACOD, a.totneto, b.sucursal, pagonumcta, pagotipcta, pagomoneda, LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2)) AS NOMBRE FROM PLAHISTORICO "
        Sql = Sql & " a INNER JOIN planillas b on (b.cia=a.cia and b.placod=a.placod and b.status!='*' and b.pagomoneda='" & Mid(cbomoneda.Text, 2, 3) & "' "
        Sql = Sql & "  and pagobanco='" & Format(cbobanco.ReturnCodigo, "00") & "') WHERE a.CIA='" & wcia & "' AND a.STATUS!='*' AND a.PROCESO='03' and "
        Sql = Sql & " month(a.fechaproceso)=" & Month(txtfecha.Value) & " and year(a.fechaproceso)=" & Year(txtfecha.Value) & " and pagonumcta!='0' ORDER BY LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2))"
       Else
        Sql = "SELECT a.PLACOD, a.totneto, b.sucursal, pagonumcta, pagotipcta, pagomoneda, LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2)) AS NOMBRE FROM PLAHISTORICO "
        Sql = Sql & " a INNER JOIN planillas b on (b.cia=a.cia and b.placod=a.placod and b.status!='*' and b.pagomoneda='" & Mid(cbomoneda.Text, 2, 3) & "' "
        Sql = Sql & "  and pagobanco='" & Format(cbobanco.ReturnCodigo, "00") & "') WHERE a.CIA='" & wcia & "' AND a.STATUS!='*' AND a.PROCESO='06' and "
        Sql = Sql & " month(a.fechaproceso)=" & Month(txtfecha.Value) & " and year(a.fechaproceso)=" & Year(txtfecha.Value) & " and pagonumcta!='0' ORDER BY LTRIM(RTRIM(b.ap_pat))+' '+LTRIM(RTRIM(ap_mat))+' '+LTRIM(RTRIM(ap_cas))+' '+LTRIM(RTRIM(nom_1))+' '+LTRIM(RTRIM(nom_2))"
End If

sumacta = 0
SUMIMPORTES = 0
Set Rs = cn.Execute(Sql)
vbgcts.Clear

Do While Not Rs.EOF
    If Len(Trim(Rs!pagonumcta & "")) > 0 Then
        sumacta = sumacta + 1 'CVar(Rs!pagonumcta)
        SUMIMPORTES = SUMIMPORTES + Rs!totneto
        
        If Rs!pagotipcta = "01" Then IDENT = 4
        
        With vbgcts
            .AddRow
            .CellDetails .Rows, 1, , , iCHCKACT, , , , 5
            .CellDetails .Rows, 2, IDENT
            .CellDetails .Rows, 3, Right("000" & Trim(Rs!SUCURSAL & ""), 3)
            .CellDetails .Rows, 4, Trim(Rs!pagonumcta & "")
            .CellDetails .Rows, 5, Format(Rs!totneto, "#0.00"), DT_RIGHT
            .CellDetails .Rows, 6, Trim(Rs!nombre)
        End With
    End If
    Rs.MoveNext
Loop
    
    vbgcts.AddRow
    ultfila = vbgcts.Rows
    vbgcts.CellDetails vbgcts.Rows, 5, Format(SUMIMPORTES, "#0.00"), DT_RIGHT
    vbgcts.CellDetails vbgcts.Rows, 6, "Total"
    
End Sub

Private Sub Cmbtipo_Click()
'VTipotrab = Funciones.fc_CodigoComboBox(Cmbtipo, 2)
VTipotrab = Funciones.fc_CodigoComboBox(Cmbtipo, 2)
Select Case VTipotrab

Case Is = "01"
    Txt_Semana.Enabled = False
    Txt_Semana.Text = " "
    UpDown1.Enabled = False
Case Is = "02"
    Txt_Semana.Enabled = True
    UpDown1.Enabled = True
End Select
End Sub

Private Sub Command1_Click()
Dim I As Long
Dim CADENAPRINT As String, CADENAPRINTDET As String
Dim IDENT As String
Dim Sql As String, Rs As ADODB.Recordset, RUTA As String
Dim count  As Integer

On Error GoTo Archivo

count = 0
For I = 1 To vbgcts.Rows - 1
    If vbgcts.CellIcon(I, 1) = iCHCKACT Then
        count = count + 1
    End If
Next

If count = 0 Then
   MsgBox "No hay Informacion para Generar Archivo", vbExclamation, "Planillas"
   Exit Sub
End If

CADENAPRINT = ""
' CODIGO DE EMPRESA
CADENAPRINT = Mid(CONS_EMPRESA, 1, (23 - 1) + 1)

'NOMBRE DE EMPRESA
Sql = "SELECT RAZSOC FROM CIA WHERE COD_CIA='" & wcia & "' AND STATUS!='*'"
Set Rs = cn.Execute(Sql)
If Not Rs.EOF Then
    CADENAPRINT = CADENAPRINT & Mid(Trim(Rs(0)) & Space((44 - 24) + 1), 1, (44 - 24) + 1)
Else
    CADENAPRINT = CADENAPRINT & Mid(Space((44 - 24) + 1), 1, (44 - 24) + 1)
End If

'SUMA DE CUENTAS
CADENAPRINT = CADENAPRINT & Right("000000000000000" & CStr(CONS_CUENTA_DET + CStr(sumacta)), 15)

'moneda
CADENAPRINT = CADENAPRINT & Mid(cbomoneda.Text, 2, 2)

'SUMA DE LOS IMPORTES
CADENAPRINT = CADENAPRINT & Right("000000000000000" & CStr(SUMIMPORTES * 100), 15)

'DIA MES
CADENAPRINT = CADENAPRINT & Format(Day(txtfecha.Value), "00") & Format(Month(txtfecha.Value), "00")


RUTA$ = App.path & "\REPORTS\" & "PAGHAB.DAT"
If Dir(RUTA$) <> "" Then
    Kill RUTA$
End If

Open RUTA$ For Output As #1
Print #1, CADENAPRINT

For I = 1 To vbgcts.Rows - 1
    If vbgcts.CellIcon(I, 1) = iCHCKACT Then
        CADENAPRINTDET = ""
        
        'IDEN TIPO DE CUENTA
        CADENAPRINTDET = Trim(vbgcts.CellText(I, 2))
        'CAMPO FIJO
        CADENAPRINTDET = CADENAPRINTDET & CONS_FIJODET
        'SUCURSAL
        CADENAPRINTDET = CADENAPRINTDET & Right("000" & Trim(vbgcts.CellText(I, 3)), 3)
        'NRO CTA
        CADENAPRINTDET = CADENAPRINTDET & Right("00000000000000" & Trim(vbgcts.CellText(I, 4)), 13)
        'NOMBRE DEL TRABAJADOR
        CADENAPRINTDET = CADENAPRINTDET & Left(Trim(vbgcts.CellText(I, 6)) & Space(36), 34) & "  "
        'MONEDA
        CADENAPRINTDET = CADENAPRINTDET & Mid(cbomoneda.Text, 2, 2)
        'NETO
        CADENAPRINTDET = CADENAPRINTDET & Right("00000000000000" & (vbgcts.CellText(I, 5) * 100), 15)
        
        Print #1, CADENAPRINTDET
    End If
Next
    
Close #1



MsgBox "Archivo se Genero Correctamente" & Chr(13) & RUTA, vbInformation, "Planillas"
RUTA$ = App.path & "\REPORTS\"
Call Imprime_Txt("PAGHAB.DAT", RUTA$)
Unload Me
Exit Sub
Archivo:
    MsgBox "Error Generando Archivo" & Chr(13) & ERR.Description, vbExclamation, "Planillas"
    If Dir(RUTA$) <> "" Then
        Kill RUTA$
    End If
   
End Sub

'Private Sub Form_DblClick()
'Procesar
'End Sub

Private Sub Form_Load()
Dim Sql As String
cbobanco.NameTab = "maestros_2"
cbobanco.NameCod = "cod_maestro2"
cbobanco.NameDesc = "descrip"
cbobanco.Filtro = " right(ciamaestro,3)='007' and status!='*' "
cbobanco.OrderBy = "cod_maestro2"
cbobanco.conexion = cn
cbobanco.Execute

cbomoneda.NameTab = "maestros_2"
cbomoneda.NameCod = "cod_maestro2"
cbomoneda.NameDesc = "flag1"
cbomoneda.Filtro = " right(ciamaestro,3)='006' and status!='*' "
cbomoneda.conexion = cn
cbomoneda.Execute


cbo_boletas.NameTab = "maestros_2"
cbo_boletas.NameCod = "cod_maestro2"
cbo_boletas.NameDesc = "descrip"
cbo_boletas.Filtro = " right(ciamaestro,3)='078' and status!='*' "
cbo_boletas.conexion = cn
cbo_boletas.Execute

txtfecha.Day = Day(Now)
txtfecha.Month = Month(Now)
txtfecha.Year = Year(Now)

Cmbtipo.Clear

   wciamae = Determina_Maestro("01055")
   Sql$ = "Select * from maestros_2 where flag1 IN ('02','04') and status<>'*'"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Rs.RecordCount > 0 Then Rs.MoveFirst
   Do While Not Rs.EOF
      Cmbtipo.AddItem Rs!descrip
      Cmbtipo.ItemData(Cmbtipo.NewIndex) = Trim(Rs!Cod_maestro2)
      Rs.MoveNext
   Loop
   Rs.Close
   If Cmbtipo.ListCount >= 0 Then Cmbtipo.ListIndex = 0


InicializaGrilla

End Sub


Private Sub InicializaGrilla()
With vbgcts
    .Redraw = False
    
    .Gridlines = True
    .GridLineMode = ecgGridFillControl
        
    .HighlightSelectedIcons = False
    .RowMode = True
    .Editable = True
    .SingleClickEdit = True
    
    .StretchLastColumnToFit = True
    
    .ScrollBarStyle = ecgSbrFlat
    .ImageList = vbalImageList1
    
    .AddColumn "chk", "", ecgHdrTextALignCentre, , 25
    .AddColumn "Iden", "", ecgHdrTextALignCentre, , , False
    .AddColumn "Sucursal", "", ecgHdrTextALignCentre, , , False
    .AddColumn "NroCta", "Nro Cta", ecgHdrTextALignCentre, , 90, True
    .AddColumn "Neto", "Importe", ecgHdrTextALignCentre, , , True
    .AddColumn "Nombre", "Nombre", ecgHdrTextALignCentre, , , True
    
    .SetHeaders
    
    .Redraw = True
End With
End Sub

Private Sub Txt_Semana_Change()
Procesa_Seteo_Boleta
End Sub

Private Sub UpDown1_DownClick()
If Txt_Semana.Text = "" Then Txt_Semana.Text = "0"
If Txt_Semana.Text > 0 Then Txt_Semana.Text = Txt_Semana - 1
End Sub

Private Sub UpDown1_UpClick()
If Txt_Semana.Text = "" Then Txt_Semana.Text = "0"
Txt_Semana.Text = Txt_Semana + 1
End Sub


Private Sub vbgcts_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
vbgcts.Redraw = False
If vbgcts.Rows > 0 Then
    If lCol = 1 Then
        bCancel = True
        If vbgcts.CellIcon(lRow, lCol) = iCHCKINAC Then
            vbgcts.CellIcon(lRow, lCol) = iCHCKACT
            SUMIMPORTES = SUMIMPORTES + CCur(vbgcts.CellText(lRow, 5))
            vbgcts.CellText(ultfila, 5) = SUMIMPORTES
            sumacta = sumacta + CVar(vbgcts.CellText(lRow, 4))
        Else
            vbgcts.CellIcon(lRow, lCol) = iCHCKINAC
            SUMIMPORTES = SUMIMPORTES - CCur(vbgcts.CellText(lRow, 5))
            vbgcts.CellText(ultfila, 5) = SUMIMPORTES
            sumacta = sumacta - CVar(vbgcts.CellText(lRow, 4))
        End If
    End If
End If
vbgcts.Redraw = True

End Sub

Public Sub Procesa_Seteo_Boleta()
Dim Sql As String
If Trim(Txt_Semana.Text) <> "" Then
Sql$ = "select fechaf from plasemanas where cia='" & wcia & "' and ano='" & Format(txtfecha.Year, "0000") & "' and semana='" & Format(Trim(Txt_Semana.Text), "00") & "'  and status<>'*'"
 
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   'Cmbdel.Value = Format(Rs!fechai, "dd/mm/yyyy")
   txtfecha.Value = Format(Rs!fechaf, "dd/mm/yyyy")
End If

    If Rs.State = 1 Then Rs.Close
    End If

End Sub


