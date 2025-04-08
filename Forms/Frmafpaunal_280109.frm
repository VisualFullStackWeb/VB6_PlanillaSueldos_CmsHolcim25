VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmafpaunal 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIQUIDACION ANUAL DE APORTES DE AFP"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "Frmafpaunal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox Cmbafp 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
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
   Begin MSForms.SpinButton SpinButton1 
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   255
      Size            =   "450;564"
   End
   Begin VB.Label Label3 
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
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Afp"
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
      TabIndex        =   3
      Top             =   720
      Width           =   300
   End
End
Attribute VB_Name = "Frmafpaunal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlinea As Integer
Dim vAfp As String
Dim mciadir As String
Dim mciatlf As String


Private Sub Cmbafp_Click()
vAfp = fc_CodigoComboBox(Cmbafp, 2)
Procesa_Afp_Anual
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01069", "", Cmbafp)
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5295
Me.Height = 1605
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Txtano.Text = Format(Year(Date), "0000")
End Sub
Public Sub Procesa_Afp_Anual()
Dim rs2 As ADODB.Recordset
Dim mcad As String
Dim mItem As Integer
Dim mtot As Currency
Dim mtot1 As Currency
Dim mtot2 As Currency
Dim mtot3 As Currency
Dim mtot4 As Currency
Dim mtot5 As Currency
Dim mtot6 As Currency
Dim mtot7 As Currency
Dim mtot8 As Currency
Dim mtot9 As Currency
Dim mtot10 As Currency
Dim mtot11 As Currency
Dim mtot12 As Currency
Dim mtot13 As Currency

If Cmbafp.ListIndex < 0 Then MsgBox "Debe Seleccionar AFP", vbInformation, "AFP": Exit Sub
mciadir = ""
Sql$ = "select direcc,nro,dpto from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then mciadir = Trim(rs!direcc) & " " & Trim(rs!NRO)
If rs.State = 1 Then rs.Close
Sql$ = "select telef from telef_cia where cod_cia='" & wcia & "'"
If (fAbrRst(rs, Sql$)) Then mciatlf = Trim(rs!telef)
If rs.State = 1 Then rs.Close

If vAfp <> "01" Then

'Sql$ = "select placod,numafp," _
'     & "sum( case when month(fechaproceso)=1 then d11 else 0.00 end)  as m01," _
'     & "sum( case when month(fechaproceso)=2 then d11 else 0.00 end)  as m02," _
'     & "sum( case when month(fechaproceso)=3 then d11 else 0.00 end)  as m03," _
'     & "sum( case when month(fechaproceso)=4 then d11 else 0.00 end)  as m04," _
'     & "sum( case when month(fechaproceso)=5 then d11 else 0.00 end)  as m05," _
'     & "sum( case when month(fechaproceso)=6 then d11 else 0.00 end)  as m06," _
'     & "sum( case when month(fechaproceso)=7 then d11 else 0.00 end)  as m07," _
'     & "sum( case when month(fechaproceso)=8 then d11 else 0.00 end)  as m08," _
'     & "sum( case when month(fechaproceso)=9 then d11 else 0.00 end)  as m09," _
'     & "sum( case when month(fechaproceso)=10 then d11 else 0.00 end)  as m10," _
'     & "sum( case when month(fechaproceso)=11 then d11 else 0.00 end)  as m11," _
'     & "sum( case when month(fechaproceso)=12 then d11 else 0.00 end)  as m12 " _
'     & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and codafp='" & vAfp & "' and status<>'*' " _
'     & "group by placod,numafp"
     
Sql$ = "select b.* from (select a.*, a.m01+a.m02+a.m03+a.m04+a.m05+a.m06+a.m07+a.m08+a.m09+a.m10+a.m11+a.m12 tot from (select placod,numafp,sum( case when month(fechaproceso)=1 then d11 else 0.00 end)  as m01, " & _
       "sum( case when month(fechaproceso)=2 then d11 else 0.00 end)  as m02, " & _
       "sum( case when month(fechaproceso)=3 then d11 else 0.00 end)  as m03, " & _
       "sum( case when month(fechaproceso)=4 then d11 else 0.00 end)  as m04, " & _
       "SUM( case when month(fechaproceso)=5 then d11 else 0.00 end)  as m05, " & _
       "sum( case when month(fechaproceso)=6 then d11 else 0.00 end)  as m06, " & _
       "sum( case when month(fechaproceso)=7 then d11 else 0.00 end)  as m07, " & _
       "sum( case when month(fechaproceso)=8 then d11 else 0.00 end)  as m08, " & _
       "sum( case when month(fechaproceso)=9 then d11 else 0.00 end)  as m09, " & _
       "sum( case when month(fechaproceso)=10 then d11 else 0.00 end)  as m10, " & _
       "sum( case when month(fechaproceso)=11 then d11 else 0.00 end)  as m11, " & _
       "sum( case when month(fechaproceso)=12 then d11 else 0.00 end)  as m12 " & _
       "from plahistorico " & _
       "where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and codafp='" & vAfp & "' and status<>'*'group by placod,numafp) a ) b where tot<>0"
     
     
 Else
 
' Sql$ = "select placod,numafp," _
'     & "sum( case when month(fechaproceso)=1 then d04 else 0.00 end)  as m01," _
'     & "sum( case when month(fechaproceso)=2 then d04 else 0.00 end)  as m02," _
'     & "sum( case when month(fechaproceso)=3 then d04 else 0.00 end)  as m03," _
'     & "sum( case when month(fechaproceso)=4 then d04 else 0.00 end)  as m04," _
'     & "sum( case when month(fechaproceso)=5 then d04 else 0.00 end)  as m05," _
'     & "sum( case when month(fechaproceso)=6 then d04 else 0.00 end)  as m06," _
'     & "sum( case when month(fechaproceso)=7 then d04 else 0.00 end)  as m07," _
'     & "sum( case when month(fechaproceso)=8 then d04 else 0.00 end)  as m08," _
'     & "sum( case when month(fechaproceso)=9 then d04 else 0.00 end)  as m09," _
'     & "sum( case when month(fechaproceso)=10 then d04 else 0.00 end)  as m10," _
'     & "sum( case when month(fechaproceso)=11 then d04 else 0.00 end)  as m11," _
'     & "sum( case when month(fechaproceso)=12 then d04 else 0.00 end)  as m12 " _
'     & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and codafp='" & vAfp & "' and status<>'*' " _
'     & "group by placod,numafp"
     
Sql$ = "select b.* from (select a.*, a.m01+a.m02+a.m03+a.m04+a.m05+a.m06+a.m07+a.m08+a.m09+a.m10+a.m11+a.m12 tot from (select placod,numafp,sum( case when month(fechaproceso)=1 then d11 else 0.00 end)  as m01, " & _
       "sum( case when month(fechaproceso)=2 then d11 else 0.00 end)  as m02, " & _
       "sum( case when month(fechaproceso)=3 then d11 else 0.00 end)  as m03, " & _
       "sum( case when month(fechaproceso)=4 then d11 else 0.00 end)  as m04, " & _
       "SUM( case when month(fechaproceso)=5 then d11 else 0.00 end)  as m05, " & _
       "sum( case when month(fechaproceso)=6 then d11 else 0.00 end)  as m06, " & _
       "sum( case when month(fechaproceso)=7 then d11 else 0.00 end)  as m07, " & _
       "sum( case when month(fechaproceso)=8 then d11 else 0.00 end)  as m08, " & _
       "sum( case when month(fechaproceso)=9 then d11 else 0.00 end)  as m09, " & _
       "sum( case when month(fechaproceso)=10 then d11 else 0.00 end)  as m10, " & _
       "sum( case when month(fechaproceso)=11 then d11 else 0.00 end)  as m11, " & _
       "sum( case when month(fechaproceso)=12 then d11 else 0.00 end)  as m12 " & _
       "from plahistorico " & _
       "where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and codafp='" & vAfp & "' and status<>'*'group by placod,numafp) a ) b where tot<>0"
     
 End If
     
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Existen Datos Registrados", vbInformation, "Liquidacion Anual de Afp": Exit Sub
rs.MoveFirst
mItem = 1
RUTA$ = App.Path & "\REPORTS\" & "Afpanual.txt"
Open RUTA$ For Output As #1
mtot = 0: mtot1 = 0: mtot2 = 0: mtot3 = 0: mtot4 = 0: mtot5 = 0: mtot6 = 0: mtot7 = 0
mtot8 = 0: mtot9 = 0: mtot10 = 0: mtot11 = 0: mtot12 = 0: mtot13 = 0:
Print #1, LetraChica
Call Cabeza_Afp_Anual
Do While Not rs.EOF
   Sql$ = "select ap_pat,ap_mat,nom_1,nom_2,fcese from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
   If Not (fAbrRst(rs2, Sql$)) Then
      MsgBox "No se Encuentra Registrado el Codigo de Trabajador", vbCritical, "Codigo => " & rs!PLACOD
      If rs.State = 1 Then rs.Close
      If rs2.State = 1 Then rs2.Close
      Exit Sub
   End If
   mcad = ""
   
   'MODIFICADO EL 08/01/2008
   NUMAFP = IIf(IsNull(rs!NUMAFP), "", rs!NUMAFP)
   mcad = fCadNum(mItem, "###") & " " & lentexto(12, Left(NUMAFP, 12)) & " " & lentexto(17, Left(rs2!ap_pat, 17)) & " " & lentexto(17, Left(rs2!ap_mat, 17)) & " " & lentexto(17, Left(Trim(rs2!nom_1) & " " & Trim(rs2!nom_2), 17))
      
   For I = 2 To 13
   
       If rs(I) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & Space(1) & fCadNum(rs(I), "###,###.00")
       End If
       mtot = mtot + (rs(I))
       Select Case I
              Case Is = 2
                    mtot1 = mtot1 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 3
                    mtot2 = mtot2 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 4
                    mtot3 = mtot3 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 5
                    mtot4 = mtot4 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 6
                    mtot5 = mtot5 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 7
                    mtot6 = mtot6 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 8
                    mtot7 = mtot7 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 9
                    mtot8 = mtot8 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 10
                    mtot9 = mtot9 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 11
                    mtot10 = mtot10 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 12
                    mtot11 = mtot11 + rs(I)
                    mtot13 = mtot13 + rs(I)
              Case Is = 13
                    mtot12 = mtot12 + rs(I)
                    mtot13 = mtot13 + rs(I)
         End Select
   Next I
   mcad = mcad & Space(1) & IIf(mtot = 0, Space(11), fCadNum(mtot, "###,###.00")) & Space(2) & Format(rs2!fcese, "dd/mm/yyyy")
   Print #1, mcad
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
   mItem = mItem + 1
   mlinea = mlinea + 1
   If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Afp_Anual
  mtot = 0
Loop
If rs.State = 1 Then rs.Close

Print #1, Space(70) & String(144, "-")

mcad = Space(71) & fCadNum(mtot1, "###,###.00") & Space(1) & fCadNum(mtot2, "###,###.00") & Space(1) & fCadNum(mtot3, "###,###.00") & Space(1) & fCadNum(mtot4, "###,###.00") & Space(1) & fCadNum(mtot5, "###,###.00") & Space(1)
mcad = mcad & fCadNum(mtot6, "###,###.00") & Space(1) & fCadNum(mtot7, "###,###.00") & Space(1) & fCadNum(mtot8, "###,###.00") & Space(1) & fCadNum(mtot9, "###,###.00") & Space(1) & fCadNum(mtot10, "###,###.00") & Space(1)
mcad = mcad & fCadNum(mtot11, "###,###.00") & Space(1) & fCadNum(mtot12, "###,###.00") & Space(1) & fCadNum(mtot13, "###,###.00")

Print #1, mcad
Close #1
Call Imprime_Txt("Afpanual.txt", RUTA$)
End Sub
Private Sub Cabeza_Afp_Anual()
Print #1, "LIQUIDACION ANUAL DE APORTES Y RETENCIONES PREVISIONALES" & " " & Txtano.Text
Print #1, "Empresa     " & Cmbcia.Text & Space(5) & "Ruc   :  " & wruc
Print #1, "Direccion   " & mciadir & "             Telefono  : " & mciatlf
Print #1,
Print #1, "AFP         " & Cmbafp.Text
Print #1,
Print #1, " No    CUSPP     Apellido Paterno  Apellido Materno  Nombre                  Ene        feb        Mar        Abr        May        Jun       Jul        Ago        Sep        Oct         Nov        Dic       Total    F.Cese"
Print #1, String(225, "-")
mlinea = 11
End Sub



Private Sub SpinButton1_SpinDown()
If Val(Txtano.Text) > 0 Then Txtano = Txtano - 1
End Sub

Private Sub SpinButton1_SpinUp()
If Val(Txtano) = 0 Then Txtano = "1" Else Txtano = Txtano + 1
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Procesa_Afp_Anual
End Sub
