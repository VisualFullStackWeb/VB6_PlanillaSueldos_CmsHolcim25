VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmFirmaContrato 
   Caption         =   "Firmantes de Contratos"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   Icon            =   "FrmFirmaContrato.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   5265
   Begin Threed.SSPanel SSPanel1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   2566
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TxtCargoEmpleado 
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox TxtDNIEmpleado 
         Height          =   285
         Left            =   4320
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtNombreEmpleado 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DNI"
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
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Firmante Para Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5295
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   2778
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TxtCargoObrero 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox TxtNombreObrero 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox TxtDNIObrero 
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Firmante Para Obreros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNI"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   345
      End
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   3600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "FrmFirmaContrato.frx":030A
   End
End
Attribute VB_Name = "FrmFirmaContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 5385
Me.Height = 4920

Sql$ = "select * from Pla_Firma_Contrato where cia='" & wcia & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If rs!TipoTrab = "01" Then
      TxtNombreEmpleado.Text = Trim(rs!nombre & "")
      TxtDNIEmpleado.Text = Trim(rs!DNI & "")
      TxtCargoEmpleado.Text = Trim(rs!Cargo & "")
   End If
   If rs!TipoTrab = "02" Then
      TxtNombreObrero.Text = Trim(rs!nombre & "")
      TxtDNIObrero.Text = Trim(rs!DNI & "")
      TxtCargoObrero.Text = Trim(rs!Cargo & "")
   End If
   rs.MoveNext
Loop
rs.Close: Set rs = Nothing
End Sub
Private Sub SSCommand7_Click()

Sql$ = "update Pla_Firma_Contrato set status='*',UserModi='" & wuser & "',FecModi=Getdate() where cia='" & wcia & "' and status<>'*'"
cn.Execute Sql$

Sql$ = "Insert Into Pla_Firma_Contrato Values('" & wcia & "','" & Trim(TxtNombreEmpleado.Text) & "','" & Trim(TxtDNIEmpleado.Text) & "','','" & wuser & "',Getdate(),'" & wuser & "',Getdate(),'01','" & Trim(TxtCargoEmpleado.Text) & "')"
cn.Execute Sql$

Sql$ = "Insert Into Pla_Firma_Contrato Values('" & wcia & "','" & Trim(TxtNombreObrero.Text) & "','" & Trim(TxtDNIObrero.Text) & "','','" & wuser & "',Getdate(),'" & wuser & "',Getdate(),'02','" & Trim(TxtCargoObrero.Text) & "')"
cn.Execute Sql$

MsgBox "Grabación Satisfactoria", vbInformation

End Sub
