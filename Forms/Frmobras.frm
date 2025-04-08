VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmobras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de obras"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   9975
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSFrame SSFrame3 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   4800
         Width           =   8175
         _Version        =   65536
         _ExtentX        =   14420
         _ExtentY        =   1296
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Txtcosto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6720
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox Cmbmoneda 
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Costo de la obra"
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
            Left            =   1080
            TabIndex        =   35
            Top             =   300
            Width           =   1410
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Left            =   4200
            TabIndex        =   30
            Top             =   240
            Width           =   585
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   8175
         _Version        =   65536
         _ExtentX        =   14420
         _ExtentY        =   4260
         _StockProps     =   14
         Caption         =   "Direccion Legal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand SSCommand1 
            Height          =   495
            Left            =   7200
            TabIndex        =   34
            Top             =   960
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   873
            _StockProps     =   78
            Picture         =   "Frmobras.frx":0000
         End
         Begin VB.TextBox Txtmtrs 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6720
            TabIndex        =   28
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Txtrefe 
            Height          =   285
            Left            =   1080
            MaxLength       =   60
            TabIndex        =   26
            Top             =   1560
            Width           =   6855
         End
         Begin VB.ComboBox Cmburb 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox Txtnro 
            Height          =   285
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   21
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Txtdir 
            Height          =   285
            Left            =   120
            MaxLength       =   60
            TabIndex        =   19
            Top             =   480
            Width           =   6975
         End
         Begin VB.Label Lblubigeo 
            Height          =   135
            Left            =   3960
            TabIndex        =   36
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LblUbica 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   2880
            TabIndex        =   33
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Metros sobre el nivel del mar"
            Height          =   195
            Left            =   4560
            TabIndex        =   27
            Top             =   1920
            Width           =   2010
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1560
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicacion"
            Height          =   195
            Left            =   2880
            TabIndex        =   24
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Urbanizacion"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No/Mz/ Lt"
            Height          =   195
            Left            =   7200
            TabIndex        =   20
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calle / Av. /Jr."
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1035
         End
      End
      Begin MSMask.MaskEdBox Txtfecf 
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txtfeci 
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txtlic 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Txtingresi 
         Height          =   285
         Left            =   4200
         MaxLength       =   60
         TabIndex        =   10
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Txtingresp 
         Height          =   285
         Left            =   240
         MaxLength       =   60
         TabIndex        =   8
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Txtcod 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Txtnombre 
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   4
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Termino"
         Height          =   195
         Left            =   4680
         TabIndex        =   15
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licencia"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingeniero Residente"
         Height          =   195
         Left            =   4200
         TabIndex        =   9
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingeniero Responsable"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Interno"
         Height          =   195
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la Obra"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Label Lblcia 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   7935
      End
   End
End
Attribute VB_Name = "Frmobras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VNUrb As String
Private Sub Cmburb_Click()
VNUrb = fc_CodigoComboBox(Cmburb, 3)
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 8625
Me.Height = 6645
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = "select * from tablas where cod_tipo='01'  and status<>'*' order by descrip"
Set rs = cn.Execute(Sql$)
Cmburb.Clear
Do While Not rs.EOF
   Cmburb.AddItem rs!DESCRIP
   Cmburb.ItemData(Cmburb.NewIndex) = Trim(rs!COD_MAESTRO2)
   rs.MoveNext
Loop
Call fc_Descrip_Maestros2_Mon("01006", "", Cmbmoneda)
For I = 0 To Cmbmoneda.ListCount - 1
    If Right(Left(Cmbmoneda.List(I), 4), 3) = wmoncont Then Cmbmoneda.ListIndex = I: Exit For
Next
End Sub
Public Sub Nueva_Obra(ciades)
If ciades <> "" Then LblCia.Caption = ciades
txtcod.Text = ""
Txtnombre = ""
Txtingresp = ""
Txtingresi = ""
Txtlic = ""
Txtfeci.Text = "__/__/____"
Txtfecf.Text = "__/__/____"
Txtdir = ""
TxtNro = ""
Cmburb.ListIndex = -1
Lblubigeo.Caption = ""
LblUbica = ""
Txtrefe = ""
Txtcosto = ""
Txtmtrs = ""
Cmbmoneda.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Frmgrdobra.Procesa_Obras
End Sub

Private Sub SSCommand1_Click()
Load FrmUbigeo
FrmUbigeo.Show
FrmUbigeo.ZOrder 0
End Sub
Public Sub Graba_Obra()
On Error GoTo ErrorTrans
Dim NroTrans As Integer
Dim Meps As String
Dim Messalud As String
Meps = ""
Messalud = ""
NroTrans = 0
If Txtnombre.Text = "" Then MsgBox "Debe Ingresar Nombre de la Obra", vbInformation, "Mantenimiento de Obras": Txtnombre.SetFocus: Exit Sub

Mgrab = MsgBox("Seguro de Grabar Obra", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub

Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
NroTrans = 1

If IsNumeric(Txtcosto.Text) = False Then Txtcosto.Text = "0.00"
If IsNumeric(Txtmtrs.Text) = False Then Txtmtrs.Text = "0.00"

If Trim(txtcod.Text) = "" Then
   Sql$ = "select max(cod_obra) from plaobras where cod_cia='" & wcia & "'"
'   Debug.Print SQL$
   Set rs = cn.Execute(Sql$)
   If (Funciones.fAbrRst(rs, Sql$) And Not IsNull(rs(0))) Then
      txtcod.Text = Format(rs(0) + 1, "00000000")
   Else
      txtcod.Text = "00000001"
   End If
   If rs.State = 1 Then rs.Close
End If
Sql$ = "update plaobras set status='*' where cod_cia='" & wcia & "' and cod_obra='" & Trim(txtcod.Text) & "'"
cn.Execute Sql$
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "INSERT INTO plaobras values('" & wcia & "','" & Trim(txtcod.Text) & "','" & Trim(Txtnombre.Text) & "','" & Txtingresp.Text & "','" & Txtingresi.Text & "','" & Trim(Txtlic.Text) & "', " _
    & "'" & IIf(Txtfeci.Text = "__/__/____", Null, Format(Txtfeci, FormatFecha)) & "', " _
    & "'" & IIf(Txtfecf.Text = "__/__/____", Null, Format(Txtfecf, FormatFecha)) & "', " _
    & "'" & Trim(Txtdir.Text) & "','" & Trim(TxtNro.Text) & "','" & VNUrb & "','" & Lblubigeo.Caption & "','" & Trim(Txtrefe.Text) & "', " _
    & "'" & Right(Left(Cmbmoneda.List(I), 4), 3) & "'," & CCur(Txtcosto.Text) & ",''," & FechaSys & ",'" & wuser & "'," & CCur(Txtmtrs.Text) & ")"
cn.Execute Sql$
cn.CommitTrans

MsgBox "Operacion Realizada Exitosamente", vbInformation, ""
Screen.MousePointer = vbDefault
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox ERR.Description, vbCritical, Me.Caption

End Sub
Public Sub Carga_Obra(cianom As String, codigo As String)
txtcod.Text = codigo
Sql$ = " SELECT c.*,u.dpto as dp,prov,dist,pais " _
     & " FROM plaobras c LEFT JOIN ubigeos u ON u.cod_ubi=c.cod_ubi" _
     & " WHERE c. cod_cia='" & wcia & "' AND c. cod_obra='" & codigo & "' and c.status<>'*'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
rs.MoveFirst
          
LblCia.Caption = cianom
Txtnombre = rs!DESCRIP
Txtingresp = rs!ing_resp
Txtingresi = rs!ing_res
Txtlic = rs!licencia
If Not IsNull(rs!fec_inicio) Then Txtfeci.Text = rs!fec_inicio
If Not IsNull(rs!fec_fin) Then Txtfecf.Text = rs!fec_fin
Txtdir = rs!direcc
TxtNro = rs!NRO
Call rUbiIndCmbBox(Cmburb, rs!urb, "000")
Lblubigeo.Caption = rs!cod_ubi
LblUbica = rs!pais & " - " & rs!DP & " - " & rs!PROV & " - " & rs!DIST
Txtrefe = rs!referencia
Txtcosto = rs!costo
Txtmtrs = rs!msnm
For I = 0 To Cmbmoneda.ListCount - 1
    If Right(Left(Cmbmoneda.List(I), 4), 3) = rs!moneda Then Cmbmoneda.ListIndex = I: Exit For
Next

If rs.State = 1 Then rs.Close
End Sub

Private Sub Txtcosto_KeyPress(KeyAscii As Integer)
Txtcosto.Text = Txtcosto.Text + Fc_Decimals(KeyAscii)
End Sub

Private Sub Txtmtrs_KeyPress(KeyAscii As Integer)
Txtmtrs.Text = Txtmtrs.Text + Fc_Decimals(KeyAscii)
End Sub
Public Sub Elimina_Obra()
If MsgBox("Desea Eliminar Obra ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Administracion de Obras") = vbNo Then Exit Sub
Screen.MousePointer = vbArrowHourglass
Sql$ = wInicioTrans
cn.Execute Sql$
Sql$ = "update plaobras set status='*' where cod_cia='" & wcia & "' and cod_obra='" & Trim(txtcod.Text) & "'"
cn.Execute Sql$
Sql$ = wFinTrans
cn.Execute Sql$
Screen.MousePointer = vbDefault

End Sub
