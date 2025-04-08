VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmusers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Matenimiento de Usuarios «"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9690
   Icon            =   "Frmusers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2295
      Left            =   105
      TabIndex        =   4
      Top             =   4665
      Visible         =   0   'False
      Width           =   2955
      _Version        =   65536
      _ExtentX        =   5212
      _ExtentY        =   4048
      _StockProps     =   15
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox Txtuser 
         Height          =   285
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Txtclave1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1590
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txtclave2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1590
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   1590
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Cancelar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "Frmusers.frx":030A
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   360
         Left            =   150
         TabIndex        =   11
         Top             =   1680
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "Aceptar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         AutoSize        =   1
         MouseIcon       =   "Frmusers.frx":0326
         Picture         =   "Frmusers.frx":2AD8
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   2070
         Index           =   0
         Left            =   90
         Top             =   135
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   750
         TabIndex        =   10
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
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
         Left            =   870
         TabIndex        =   8
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirme Clave"
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
         Left            =   150
         TabIndex        =   6
         Top             =   1200
         Width           =   1275
      End
   End
   Begin VB.Frame Frametipo 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2250
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CheckBox chkContratos 
         BackColor       =   &H00808080&
         Caption         =   "Contratos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chkRemuneraciones 
         BackColor       =   &H00808080&
         Caption         =   "Remuneraciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox Cmbtipopla 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox Chckadmin 
         BackColor       =   &H00808080&
         Caption         =   "Administrador"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   2070
         Index           =   1
         Left            =   60
         Top             =   45
         Width           =   1830
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Planilla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Lbluser 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7095
      Left            =   4365
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   12515
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmusers.frx":2AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmusers.frx":308E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7095
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   12515
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnufile 
      Caption         =   "Usuarios"
      Begin VB.Menu mnunuevo 
         Caption         =   "Nuevo Usuario"
      End
      Begin VB.Menu Mnuelimina 
         Caption         =   "Eliminar Usuario"
      End
   End
End
Attribute VB_Name = "Frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Eliminar_Usuario()
On Error GoTo ErrorTrans
Dim NroTrans As Integer
NroTrans = 0
If MsgBox("Desea Eliminar Usuario ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    cn.BeginTrans
    NroTrans = 1
    Sql$ = "delete FROM users where cod_cia='" & wcia & "' and sistema='04' and name_user='" & TreeView1.SelectedItem.Text & "'"
    cn.Execute Sql$
    
    Sql$ = "delete FROM users_menu where sistema='04' and cia='" & wcia & "' and name_user='" & TreeView1.SelectedItem.Text & "'"
    cn.Execute Sql$
    
    Select Case StrConv(gsAdminDB, 1)
           Case Is = "MYSQL"
                Sql$ = "delete from mysql.user where user='" & TreeView1.SelectedItem.Text & "'"
                cn.Execute Sql$
                Sql$ = "delete from sysusers where uid='04' and name='" & TreeView1.SelectedItem.Text & "'"
               cn.Execute Sql$
    End Select
    
    cn.CommitTrans
    Carga_treeview
End If
Screen.MousePointer = vbDefault
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox Err.Description, vbInformation, Me.Caption
Screen.MousePointer = vbDefault
End Sub

Public Sub Nuevo_User()
Chckadmin.Value = 0
Frametipo.Visible = False
SSPanel1.Visible = True
Txtuser.SetFocus
End Sub
Private Sub Carga_treeview()
On Error GoTo CORRIGE
TreeView1.Nodes.Clear
TreeView1.Nodes.Add = "Usuarios"
TreeView1.Nodes(1).Image = 2
Sql$ = "select a.*,b.name from users a,sysusers b where a.sistema='04' and a.name_user=b.name  and a.status<>'*' and cod_cia='" & wcia & "' ORDER BY NAME"
 
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   Do While Not Rs.EOF
      TreeView1.Nodes.Add(1, tvwChild, , , 1) = Rs!Name
      Rs.MoveNext
   Loop
End If
TreeView1.Nodes(1).Expanded = True
If Rs.State = 1 Then Rs.Close
Exit Sub
CORRIGE:
   MsgBox "Error :" & Err.Description, vbCritical, Me.Caption
End Sub
Public Sub Grabar_users()
On Error GoTo ErrorTrans
Dim NroTrans As Integer
Dim mmenu As String
Dim mAdmin As String
Dim mTipo As String
NroTrans = 0
mAdmin = "": mTipo = ""
If Chckadmin.Value = 1 Then mAdmin = "1"
If Cmbtipopla.Text = "TOTAL" Then
   mTipo = "99"
Else
   mTipo = fc_CodigoComboBox(Cmbtipopla, 2)
End If
If MsgBox("Desea Grabar el Seteo de Usuario ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
    
Else
    cn.BeginTrans
    NroTrans = 1
    Sql$ = "delete from users_menu where sistema='04' and status<>'*' and cia='" & wcia & "' and name_user='" & Trim(TreeView1.SelectedItem.Text) & "'"
    cn.Execute Sql$
    For I = 1 To ListView1.ListItems.count
        mmenu = ListView1.ListItems(I).SubItems(1)
        If ListView1.ListItems.Item(I).Checked = True Then
           Sql$ = "insert users_menu values('" & TreeView1.SelectedItem.Text & "','" & mmenu & "','04','','" & wcia & "','','','','','')"
           cn.Execute Sql$
        End If
    Next
   
    Sql$ = "update users set admin='" & mAdmin & "',tipopla='" & mTipo
    Sql$ = Sql$ & "',autorizar_remuneracion=" & chkRemuneraciones.Value
    Sql$ = Sql$ & ",autorizar_contrato=" & chkContratos.Value
    Sql$ = Sql$ & " where sistema='04' and status<>'*' and cod_cia='" & wcia & "' and name_user='" & Lbluser.Caption & "'"
    cn.Execute Sql$
    cn.CommitTrans
    
    MsgBox "Se guardaron los datos correctamente", vbInformation, Me.Caption
End If
Screen.MousePointer = vbDefault

Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox Err.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault

End Sub
Private Sub Carga_Menu(User)

Dim MiObjeto, MiColección, Control
Dim m As Integer
Dim mItem As Integer

Dim mspace As Integer

Frametipo.Visible = True
Cmbtipopla.ListIndex = -1
Chckadmin.Value = 0

Sql$ = "select * from users where sistema='04' and status<>'*' and cod_cia='" & wcia & "' and name_user='" & User & "'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   If Rs!admin = "1" Then Chckadmin.Value = 1
   chkRemuneraciones.Value = IIf(Rs!autorizar_remuneracion = True, 1, 0)
   chkContratos.Value = IIf(Rs!autorizar_contrato = True, 1, 0)
   
   If Rs!TIPOPLA = "99" Then
      Cmbtipopla.ListIndex = Cmbtipopla.ListCount - 1
   Else
      Call rUbiIndCmbBox(Cmbtipopla, Rs!TIPOPLA & "", "00")
   End If
   
End If
Rs.Close

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Add , , "MENU", ListView1.Width * (95 / 100)
ListView1.ColumnHeaders.Add , , "NOMBRE", ListView1.Width * (2 / 100)
ListView1.ColumnHeaders.Add , , "ITEM", ListView1.Width * (2 / 100)
ListView1.ColumnHeaders.Add , , "CABEZA", ListView1.Width * (1 / 100)

m = 1
mItem = 1
For Each MiObjeto In MDIplared
    If TypeOf MiObjeto Is Menu Then
       If MiObjeto.Caption <> "-" Then
            'magl
            
           Sql$ = "insert into plamenus values ('" & MiObjeto.Caption & "','" & MiObjeto.Tag & "','" & MiObjeto.Name & "','" & mItem & "','" & m & "')"
           cn.CursorLocation = adUseClient
           Set rs2 = New ADODB.Recordset
           Set rs2 = cn.Execute(Sql$, 64)
          
          ',gl
          ListView1.ListItems.Add = MiObjeto.Caption
          Select Case MiObjeto.Tag
                 Case Is = "M1"
                      mspace = 1
                      ListView1.ListItems.Item(m).Bold = True
                      ListView1.ListItems.Item(m).ForeColor = 255
                      ListView1.ListItems(m).SubItems(3) = "S"
                      mItem = mItem + 1
                  Case Is = "M2"
                      mspace = 3
                      ListView1.ListItems.Item(m).Text = "---" & ListView1.ListItems.Item(m).Text
                      ListView1.ListItems.Item(m).Bold = True
                      ListView1.ListItems.Item(m).ForeColor = 16711680
                  Case Is = "M3"
                      mspace = 7
                      ListView1.ListItems.Item(m).Text = "------" & ListView1.ListItems.Item(m).Text
                      ListView1.ListItems.Item(m).Bold = True
                      ListView1.ListItems.Item(m).ForeColor = 32768
                  Case Else
                      ListView1.ListItems.Item(m).Text = Space(1) & Space(mspace) & ListView1.ListItems.Item(m).Text
          End Select
          ListView1.ListItems(m).SubItems(1) = MiObjeto.Name
          ListView1.ListItems(m).SubItems(2) = mItem
          Sql$ = "select * from users_menu where sistema='04' and status<>'*' and cia='" & wcia & "' and name_user='" & User & "' and name_menu='" & MiObjeto.Name & "'"
          cn.CursorLocation = adUseClient
          Set Rs = New ADODB.Recordset
          Set Rs = cn.Execute(Sql$, 64)
          If Rs.RecordCount > 0 Then
             ListView1.ListItems.Item(m).Checked = True
          Else
             ListView1.ListItems.Item(m).Checked = False
          End If

          
          m = m + 1
       End If
    End If
Next
Set MiObjeto = Nothing
Set MiColección = Nothing
If Rs.State = 1 Then Rs.Close
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7815
Me.Width = 9780
ListView1.Height = 7095
TreeView1.Height = 7095
Call Carga_treeview
Call fc_Descrip_Maestros2("01055", "", Cmbtipopla)
Cmbtipopla.AddItem "TOTAL"
Cmbtipopla.Enabled = True
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim I, FIND As Integer
If Mid(ListView1.ListItems(Item.Index).Text, 1, 1) = " " Or Mid(ListView1.ListItems(Item.Index).Text, 1, 1) = "-" Then Exit Sub
Item.Selected = True
FIND = ListView1.ListItems(Item.Index).SubItems(2)
If FIND > 0 Then
    If Item.Checked Then
        Item.Checked = True
        Call MarcarSubNivel(ListView1.ListItems(Item.Index).SubItems(1), True, FIND)
    Else
        Item.Checked = False
        Call MarcarSubNivel(ListView1.ListItems(Item.Index).SubItems(1), False, FIND)
    End If
End If
End Sub

Private Sub Mnuelimina_Click()
    Call Eliminar_Usuario
End Sub

Private Sub mnunuevo_Click()
    Call Nuevo_User
End Sub

Private Sub SSCommand1_Click()
On Error GoTo ErrorTrans
Dim NroTrans As Integer
NroTrans = 0
If Txtuser.Text = "" Then MsgBox "Ingrese Usuario", vbInformation, "Apertura de Usuario": Exit Sub
Txtuser.Text = UCase(Txtuser.Text)
Sql$ = "select * from sysusers where name='" & Trim(Txtuser.Text) & "'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)

cn.BeginTrans
NroTrans = 1
   
If Rs.RecordCount > 0 Then
Else

   Select Case StrConv(gsAdminDB, 1)
       Case Is = "MYSQL"
            Sql$ = "insert mysql.user values('%','" & Txtuser.Text & "','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','','','','',0,0,0)"
            cn.Execute Sql$
            Sql$ = "Flush PRIVILEGES"
            cn.Execute Sql$
            Sql$ = "Insert into sysusers values('04','','" & Txtuser.Text & "')"
            cn.Execute Sql$
            
       Case Is = "SQL SERVER"
       
'            Sql$ = "SP_ADDLOGIN '" & Trim(Txtuser.Text) & "','" & Trim(Txtclave1.Text) & "','" & Trim(cn.DefaultDatabase) & "'"
'            cn.Execute Sql$
'            Sql$ = "SP_ADDUSER '" & Trim(Txtuser.Text) & "','" & Trim(Txtuser.Text) & "'"
'            'SQL$ = "SP_ADDUSER '" & Trim(Txtuser.Text) & "','" & Trim(Txtuser.Text) & "','USUARIOS'"
           
           'IMPLEMENTACION GALLOS
            Sql$ = "spCrearUsuario '" & Trim(Txtuser.Text) & "','" & Trim(Txtclave1.Text) & "','" & Trim(cn.DefaultDatabase) & "'"

            cn.Execute Sql$
   End Select
End If

'If StrConv(gsAdminDB, 1) <> "MYSQL" Then
'   Sql$ = wInicioTrans
'   cn.Execute Sql$
'End If
Sql$ = "Insert into users(cod_cia, login, name_user, password, status, sistema, admin, " & _
"tipopla, encargado, cargo, email, permisos, segunclave, iniciales ,autorizar_remuneracion,autorizar_contrato) values('" & wcia & "','" & _
Txtuser.Text & "','" & Txtuser.Text & "','" & Txtclave1.Text & "','','04','',''" & _
",'','','','','',''," & chkRemuneraciones.Value & "," & chkContratos.Value & ")"

cn.Execute Sql$

cn.CommitTrans

SSPanel1.Visible = False

Carga_treeview
Frametipo.Visible = True
If Rs.State = 1 Then Rs.Close

Exit Sub
ErrorTrans:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub SSCommand2_Click()
Txtuser.Text = ""
Txtclave1.Text = ""
Txtclave2.Text = ""
SSPanel1.Visible = False
Frametipo.Visible = True
End Sub

Private Sub TreeView1_Click()
Dim wUsuario As String
If TreeView1.SelectedItem.Children = 0 Then
   wUsuario = TreeView1.SelectedItem.Text
   Lbluser = wUsuario
   Lbluser.Refresh
   Carga_Menu (wUsuario)
Else
   ListView1.ListItems.Clear
   Lbluser = ""
End If
End Sub
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnufile
End If
End Sub

Private Sub Txtclave1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtclave2.SetFocus
End Sub

Private Sub Txtclave2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtclave2_LostFocus
End Sub

Private Sub Txtclave2_LostFocus()
If Txtclave1.Text <> Txtclave2.Text Then
   MsgBox "La confirmacion de la Clave no coincide con la Clave", vbCritical, "Apertura de Usuario"
   Txtclave1.Text = ""
   Txtclave2.Text = ""
   Txtclave2.SetFocus
Else
   SSCommand1.SetFocus
End If
End Sub

Private Sub Txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtclave1.SetFocus
End Sub

Private Sub Txtuser_LostFocus()

Sql$ = "select * from users where cod_cia='" & _
wcia & "' and sistema='04' and name_user='" & _
Trim(Txtuser.Text) & "' and status<>'*'"

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   MsgBox "Usuario Ya Registrado", vbExclamation, "Apertura de Usuario"
   Txtuser.SetFocus
   Exit Sub
End If
If Rs.State = 1 Then Rs.Close
Lbluser = Txtuser
End Sub
Public Sub MarcarSubNivel(ByVal PCadMenu As String, ByVal pStatus As Boolean, mItem As Integer)
Dim I As Integer
    If ListView1.Visible Then
        For I = 1 To ListView1.ListItems.count
            If ListView1.ListItems(I).SubItems(2) = mItem Then
               ListView1.ListItems.Item(I).Checked = pStatus
            End If
        Next
    End If
End Sub

