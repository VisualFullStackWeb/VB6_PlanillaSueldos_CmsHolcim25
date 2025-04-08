VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCtsFaltaEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FrmCtsFalta"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5940
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5655
      Begin MSComCtl2.DTPicker dtfechafalta 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63635457
         CurrentDate     =   40121
      End
      Begin VB.TextBox txtpersonal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtplacod 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Personal"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Falta:"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ComboBox Cmbcia 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmCtsFaltaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Nuevo()
    Me.txtplacod.Text = ""
    Me.txtpersonal.Text = ""
    Me.dtfechafalta.Value = Now
End Sub
Public Sub Grabar()

     Sql$ = "SELECT placod from cts_falta WHERE CIA='" & wcia & "' and placod='" & Me.txtplacod.Text & "'  AND fecha_falta = convert(datetime,'" & Format(dtfechafalta.Value, "dd/mm/YYYY") & "',103) and status<>'*'"
     
    If fAbrRst(rs, Sql$) Then
        MsgBox "Ya ingreso el Dia Falta", vbExclamation
        Exit Sub
    Else
    
     Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
     Sql$ = Sql$ & "insert into cts_falta(cia, placod,fecha_falta,status,user_crea,fec_crea)" & _
            "values( '" & wcia & "','" & Me.txtplacod.Text & "','" & Format(dtfechafalta.Value, FormatFecha) & "','', '" & wuser & "',getdate())"
      cn.Execute Sql$
            
      If MsgBox("Registro Guardado. Desea Agregar mas Registros?", vbYesNo + vbQuestion) = vbYes Then
        Nuevo
        txtplacod.SetFocus
      Else
        
        Unload Me
      End If
    End If
End Sub

Private Sub Form_Load()
    Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
    Call rUbiIndCmbBox(Cmbcia, wcia, "00")
    Me.dtfechafalta.Value = Date
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FrmCtsFalta.dgctsfalta.Refresh
    FrmCtsFalta.carga_cts_falta
End Sub

Private Sub txtplacod_Change()
xpersonal = Me.txtplacod.Text
    Sql$ = "select top 1  rtrim(ap_pat) + ' ' + rtrim(ap_mat) + ' ' + rtrim(nom_1) + ' ' + rtrim(nom_2) as empleado " & _
    "From planillas " & _
    "where cia='" & wcia & "' and placod='" & xpersonal & "' and status<>'*'"
    
     cn.CursorLocation = adUseClient
     Set rs = New ADODB.Recordset
     Set rs = cn.Execute(Sql$, 64)
    
    If rs.RecordCount > 0 Then
        Me.txtpersonal.Text = rs!empleado
    Else
        Me.txtpersonal.Text = ""
    End If
    
    Set rs = Nothing
End Sub
