VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSistemaContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seteo Sistema Contable Roda S.A."
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9540
   Begin VB.Frame Frame2 
      Caption         =   "Configuración Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   8535
      Begin VB.ListBox lstSubDiario 
         Height          =   1230
         Left            =   3480
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ComboBox CboTipo_Trab 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSistemaContable.frx":0000
         Left            =   1605
         List            =   "frmSistemaContable.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2670
      End
      Begin MSDataGridLib.DataGrid dgConmay 
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "IdBoleta"
            Caption         =   "IdBoleta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Boleta"
            Caption         =   "Boleta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "IdSubDiario"
            Caption         =   "IdSubDiario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "SubDiario"
            Caption         =   "SubDiario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cgVoucher"
            Caption         =   "Voucher"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               Button          =   -1  'True
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1500.095
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo. Trabajador"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa Realacionada StartSoft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4935
      Begin MSDataGridLib.DataGrid dgEmpresa 
         Height          =   3015
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "AYO"
            Caption         =   "Año"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Codigo"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraSistemaContable 
      Caption         =   "Sistema Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.OptionButton optStartSoft 
         Caption         =   "Sistema Contable StartSoft"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton optRoda 
         Caption         =   "Sistema Contable RODA S.A."
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSistemaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmpresas As New ADODB.Recordset
Dim rsSeteoConmay As New ADODB.Recordset
Dim s_CodTipo As String

Private Sub Ejecuta_Llenado_Grilla()
Dim rsTemp As New ADODB.Recordset

Call Crea_RsSeteoConMay
Sql$ = "usp_Pla_SeteoConmay '" & wcia & "','" & Trim(s_CodTipo) & "'"
If (fAbrRst(rsTemp, Sql$)) Then
     rsTemp.MoveFirst
     Do While Not rsTemp.EOF
         rsSeteoConmay.AddNew
         rsSeteoConmay!IdBoleta = rsTemp!IdBoleta
         rsSeteoConmay!Boleta = rsTemp!Boleta
         rsSeteoConmay!IdSubDiario = rsTemp!IdSubDiario
         rsSeteoConmay!SubDiario = rsTemp!SubDiario
         rsSeteoConmay!cgVoucher = rsTemp!cgVoucher
         rsSeteoConmay.Update
         rsTemp.MoveNext
     Loop
     rsSeteoConmay.MoveFirst
     rsTemp.Close
     Set rsTemp = Nothing
End If
End Sub

Private Sub CboTipo_Trab_Click()
    s_CodTipo = Empty
    s_CodTipo = Trim(fc_CodigoComboBox(CboTipo_Trab, 2))
    If CboTipo_Trab.ListIndex <> 0 Then
        Call Ejecuta_Llenado_Grilla
    Else
        Call Crea_RsSeteoConMay
    End If

End Sub

Private Sub dgConmay_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 3 Then
    If Trim(dgConmay.Columns(3)) = "" Then
        dgConmay.Columns(2) = 0
    End If

End If
End Sub


Private Sub dgConmay_ButtonClick(ByVal ColIndex As Integer)
If lstSubDiario.Visible = True Then
    lstSubDiario.Visible = False
    Exit Sub
End If

Dim Y As Integer, xtop As Integer, xleft As Integer
Y = dgConmay.Row
xtop = dgConmay.Top + dgConmay.RowTop(Y) + dgConmay.RowHeight
Select Case dgConmay.Col
Case 3
xleft = dgConmay.Left + dgConmay.Columns(3).Left
       If Y < 5 Then
         lstSubDiario.Top = xtop
        Else
         lstSubDiario.Top = dgConmay.Top + dgConmay.RowTop(Y) - lstSubDiario.Height
        End If
        If ColIndex = 3 Then lstSubDiario.Left = xleft Else lstSubDiario.Left = dgConmay.Left + dgConmay.Columns(3).Left
         
        lstSubDiario.Visible = True
        lstSubDiario.SetFocus

End Select
End Sub

Private Sub Form_Load()

Me.ScaleHeight = 1695
Me.ScaleWidth = 5265
Me.Height = 1695
Me.Width = 5265
Call CargarDatos
Call Trae_Tipo_Trab(CboTipo_Trab)
Sql = "usp_Pla_ListarSubDiario '" & wcia & "'"
Call rCarListBox(Me.lstSubDiario, Sql)
End Sub

Private Sub Crea_Rs()
    If rsEmpresas.State = 1 Then rsEmpresas.Close
    rsEmpresas.Fields.Append "AYO", adInteger, , adFldIsNullable
    rsEmpresas.Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
    rsEmpresas.Open
    Set dgEmpresa.DataSource = rsEmpresas
   
    Call Crea_RsSeteoConMay
    
End Sub

Private Sub Crea_RsSeteoConMay()
    If rsSeteoConmay.State = 1 Then rsSeteoConmay.Close
    rsSeteoConmay.Fields.Append "IdBoleta", adChar, 2, adFldIsNullable
    rsSeteoConmay.Fields.Append "Boleta", adVarChar, 100, adFldIsNullable
    rsSeteoConmay.Fields.Append "IdSubDiario", adInteger, , adFldIsNullable
    rsSeteoConmay.Fields.Append "SubDiario", adVarChar, 500, adFldIsNullable
    rsSeteoConmay.Fields.Append "cgVoucher", adChar, 10, adFldIsNullable
    rsSeteoConmay.Open
    Set dgConmay.DataSource = rsSeteoConmay
    
    dgConmay.Columns(3).Button = True
End Sub

Private Sub CargarDatos()
Sql$ = "SELECT syscontable from cia where cod_cia='" & wcia & "' and status<>'*'"

If (fAbrRst(rs, Sql$)) Then
   If Trim(rs!syscontable) = "01" Then
        optRoda.Value = True
        Call optRoda_Click
   Else
        optStartSoft.Value = True
        Call optStartSoft_Click
   End If
   Call Crea_Rs
   Dim rsTemp As New ADODB.Recordset
   If optRoda.Value = True Then
        Sql$ = "SELECT * FROM EMP_RELACION where emp_id='" & wcia & "' order by ayo"
        If (fAbrRst(rsTemp, Sql$)) Then
             rsTemp.MoveFirst
             Do While Not rsTemp.EOF
                 rsEmpresas.AddNew
                 rsEmpresas!ayo = rsTemp!ayo
                 rsEmpresas!codigo = rsTemp!EMP_RELACION
                 rsEmpresas.Update
                 rsTemp.MoveNext
             Loop
             rsTemp.Close
             Set rsTemp = Nothing
        End If
    End If
End If
If rs.State = 1 Then rs.Close

End Sub

Public Sub Grabar_Sistema()
Dim NroTrans As Integer
Dim NroMensaje As Integer
Dim Anio As Integer

On Error GoTo ErrorTrans
NroTrans = 0
NroMensaje = 0
If optRoda.Value = False And optStartSoft.Value = False Then
    MsgBox "Checkear al Sistema Contable Relacionado", vbCritical, Me.Caption
    Exit Sub
End If
cn.BeginTrans
NroTrans = 1
Sql$ = "update cia set syscontable='" & IIf(Me.optRoda.Value = True, "01", "02") & "' where cod_cia='" & wcia & "'"
cn.Execute Sql$

Sql$ = "DELETE FROM EMP_RELACION where emp_id='" & wcia & "'"
cn.Execute Sql$

If optStartSoft.Value = True Then
    If rsEmpresas.RecordCount > 0 Then
        rsEmpresas.MoveFirst
        Do While Not rsEmpresas.EOF
            If IsNull(rsEmpresas!ayo) Then
                NroMensaje = 2
                GoTo ErrorTrans
            End If
            If IsNull(rsEmpresas!codigo) Then
                NroMensaje = 3
                GoTo ErrorTrans
            End If
            If (rsEmpresas!ayo) < 1900 Then
                NroMensaje = 4
                GoTo ErrorTrans
            End If
            
            If Trim(rsEmpresas!codigo) = "" Then
                NroMensaje = 5
                GoTo ErrorTrans
            End If
            Dim rsTemp1 As New ADODB.Recordset
             Sql$ = "SELECT * FROM EMP_RELACION where emp_id='" & wcia & "' AND AYO=" & Val(rsEmpresas!ayo)
            If (fAbrRst(rsTemp1, Sql$)) Then
                NroMensaje = 6
                Anio = Val(rsEmpresas!ayo)
                GoTo ErrorTrans
            End If
            rsTemp1.Close
            Set rsTemp1 = Nothing
       
            Sql$ = "insert into EMP_RELACION (EMP_ID,EMP_RELACION,AYO) VALUES('" & wcia & "','" & Trim(rsEmpresas!codigo) & "'," & Val(rsEmpresas!ayo) & ")"
            cn.Execute Sql$
            
            rsEmpresas.MoveNext
        Loop
    End If
Else
    
    If CboTipo_Trab.ListIndex <> 0 Then
        s_CodTipo = ""
        s_CodTipo = Trim(fc_CodigoComboBox(CboTipo_Trab, 2))
    
        
        Sql$ = "usp_Pla_MantenedorPlaSeteoConmay '" & wcia & "','" & Trim(s_CodTipo) & "','',0,'','" & wuser & "',2"
        cn.Execute Sql$
    
        If rsSeteoConmay.RecordCount > 0 Then
            rsSeteoConmay.MoveFirst
            Do While Not rsSeteoConmay.EOF
                If IsNull(rsSeteoConmay!IdSubDiario) = False And IsNull(rsSeteoConmay!cgVoucher) = False Then
                    If Val(rsSeteoConmay!IdSubDiario) > 0 And Len(Trim(rsSeteoConmay!cgVoucher)) = 10 Then
                        Sql$ = "usp_Pla_MantenedorPlaSeteoConmay '" & wcia & "','" & Trim(s_CodTipo) & _
                                "','" & Trim(rsSeteoConmay!IdBoleta) & "'," & Val(rsSeteoConmay!IdSubDiario) & _
                                ",'" & Trim(rsSeteoConmay!cgVoucher) & "','" & wuser & "',1"
                        cn.Execute Sql$
                    End If
                End If
                rsSeteoConmay.MoveNext
            Loop
        End If
        
        Call Ejecuta_Llenado_Grilla
    End If
End If

cn.CommitTrans
MsgBox "Se guardarón los datos satisfactoriamente", vbInformation, Me.Caption
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
If NroMensaje = 0 Then
    MsgBox ERR.Description, vbCritical, Me.Caption
ElseIf NroMensaje = 2 Then
    MsgBox "Ingrese el año"
ElseIf NroMensaje = 3 Or NroMensaje = 5 Then
    MsgBox "Ingrese el código de la empresa relacionada"
ElseIf NroMensaje = 4 Then
    MsgBox "Ingrese un año válido"
ElseIf NroMensaje = 6 Then
    MsgBox "Ya se agregó el Año:" & Anio
End If
End Sub

Private Sub lstSubDiario_Click()
If dgConmay.Col = 3 Then
    dgConmay.Columns(3) = Trim(lstSubDiario.Text)
    dgConmay.Columns(2) = fc_CodigoIntListBox(lstSubDiario)
End If
lstSubDiario.Visible = False

End Sub

Private Sub optRoda_Click()
    dgEmpresa.Visible = Not optRoda.Value
    Me.ScaleHeight = 5415
    Me.ScaleWidth = 8760
    Me.Height = 5415
    Me.Width = 8760
    Me.Frame1.Visible = False
    Me.Frame2.Visible = True
    Me.Refresh
End Sub

Private Sub optStartSoft_Click()
    dgEmpresa.Visible = optStartSoft.Value
    Me.ScaleHeight = 5415
    Me.ScaleWidth = 5175
    Me.Height = 5415
    Me.Width = 5175
    Me.Frame2.Visible = False
    Me.Frame1.Visible = True
    Me.Scale
End Sub
