VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmtipcamb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Tipo de Cambio Contable"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "Frmtipcamb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   -360
      TabIndex        =   1
      Top             =   720
      Width           =   8415
      Begin VB.ListBox lstfactor 
         Height          =   450
         ItemData        =   "Frmtipcamb.frx":030A
         Left            =   2400
         List            =   "Frmtipcamb.frx":0314
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox lstmon 
         Height          =   645
         Left            =   720
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   795
         Left            =   5720
         TabIndex        =   11
         Top             =   120
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "Fecha"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   795
         Left            =   1610
         TabIndex        =   10
         Top             =   130
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1402
         _StockProps     =   15
         Caption         =   "Destino"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   795
         Left            =   810
         TabIndex        =   9
         Top             =   130
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1393
         _ExtentY        =   1393
         _StockProps     =   15
         Caption         =   "Origen"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   390
         Left            =   4410
         TabIndex        =   8
         Top             =   525
         Width           =   1300
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "Venta Contable"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   390
         Left            =   3120
         TabIndex        =   7
         Top             =   525
         Width           =   1290
         _Version        =   65536
         _ExtentX        =   2275
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "Compra Contable"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   510
         Left            =   3120
         TabIndex        =   6
         Top             =   120
         Width           =   2600
         _Version        =   65536
         _ExtentX        =   4586
         _ExtentY        =   900
         _StockProps     =   15
         Caption         =   "Tipo de Cambio"
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GrdTipCamb 
         Height          =   2895
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   4
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "mon_ori"
            Caption         =   "Origen"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd-MM-yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "mon_des"
            Caption         =   " Destino"
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
            DataField       =   "factor"
            Caption         =   " Factor"
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
            DataField       =   "compra"
            Caption         =   "Compra Contable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "contable"
            Caption         =   "Venta Contable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fec_crea"
            Caption         =   "                  Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   1
               Button          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Button          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Button          =   -1  'True
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2415.118
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Lblfecha 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
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
         TabIndex        =   3
         Top             =   240
         Width           =   840
      End
   End
End
Attribute VB_Name = "Frmtipcamb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTipCamb As New Recordset
Public Sub Graba_TIpCamb()
If RsTipCamb.RecordCount = 0 Then MsgBox "No Existen Datos a Guardar", vbCritical, "Tipo de Cambio": Screen.MousePointer = vbDefault: Exit Sub

If RsTipCamb.RecordCount > 0 Then RsTipCamb.MoveFirst
Do While Not RsTipCamb.EOF
   If RsTipCamb!MON_ORI = "" Or IsNull(RsTipCamb!MON_ORI) Then
      MsgBox "Debe escoger una Moneda Origen hay" & Chr(13) & "una celda en blanco", vbCritical
      GrdTipCamb.SetFocus
      Exit Sub
   ElseIf RsTipCamb!MON_DES = "" Or IsNull(RsTipCamb!MON_DES) Then
      MsgBox "Debe escoger una Moneda Destino hay" & Chr(13) & "una celda en blanco", vbCritical
      GrdTipCamb.SetFocus
      Exit Sub
   ElseIf RsTipCamb!factor = "" Or IsNull(RsTipCamb!factor) Then
      MsgBox "En Factor hay" & Chr(13) & "una celda en blanco" & Chr(13) & "una celda en blanco", vbCritical
      GrdTipCamb.SetFocus
      Exit Sub
   ElseIf IsNull(RsTipCamb!compra) Or Val(RsTipCamb!compra) <= 0 Then
      MsgBox "Ingrese Tipo de Cambio mayor a 0", vbCritical
      GrdTipCamb.Col = 3
      GrdTipCamb.SetFocus
      Exit Sub
   ElseIf IsNull(RsTipCamb!contable) Or Val(RsTipCamb!contable) <= 0 Then
      MsgBox "Ingrese Tipo de Cambio mayor a 0", vbCritical
      GrdTipCamb.Col = 4
      GrdTipCamb.SetFocus
      Exit Sub
   ElseIf IsNull(RsTipCamb!fec_crea) Or RsTipCamb!fec_crea = "" Then
      MsgBox "En campo Fecha hay" & Chr(13) & "una celda en blanco" & Chr(13) & "una celda en blanco", vbCritical
      GrdTipCamb.SetFocus
      Exit Sub
      
   End If
   RsTipCamb.MoveNext
Loop

Mgrab = MsgBox("Seguro de Grabar Tipo de Cambio", vbYesNo + vbQuestion, "Tipo de Cambio")
If Mgrab <> 6 Then Screen.MousePointer = vbDefault: Exit Sub
RsTipCamb.MoveFirst

SQL$ = wInicioTrans
cn.Execute SQL$
Dim cad As String
cad = ""
cad = Day(RsTipCamb!fec_crea)
cad = " and " & wFuncdia & "(fec_crea)=" & cad
Do While Not RsTipCamb.EOF

   SQL$ = "UPDATE tipo_cambio set status='*' " _
        & " WHERE cod_cia = '" & wcia & "' " _
        & " AND Month(fec_crea) = " & Month(RsTipCamb!fec_crea) & " " _
        & " And Year(fec_crea) = " & Year(RsTipCamb!fec_crea) & " " _
        & " AND mon_ori='" & Trim(RsTipCamb!MON_ORI) & "' and mon_des='" & Trim(RsTipCamb!MON_DES) & "' and status<>'*'"
   SQL$ = SQL$ & cad
   cn.Execute SQL$

   SQL$ = " INSERT INTO tipo_cambio(cod_cia,mon_ori,mon_des,factor,venta,compra,contable,user_crea,status,fec_crea) " _
        & " VALUES ('" & wcia & "','" & Trim(RsTipCamb!MON_ORI) & "','" & Trim(RsTipCamb!MON_DES) & "' , '" & Trim(RsTipCamb!factor) & "', " _
        & " 0," & CCur(RsTipCamb!compra) & "," & CCur(RsTipCamb!contable) & ", " _
        & " '" & wuser & "',''," & FechaSys & ")"
   cn.Execute SQL$
   RsTipCamb.MoveNext
Loop

SQL$ = wFinTrans
cn.Execute SQL$

Screen.MousePointer = vbDefault
End Sub
Private Sub Cmbcia_Click()

wcia = Format(Cmbcia.ItemData(Cmbcia.ListIndex), "00")
GrdTipCamb.ClearFields
If RsTipCamb.State = 1 Then RsTipCamb.Close
RsTipCamb.Fields.Append "mon_ori", adChar, 4, adFldIsNullable
RsTipCamb.Fields.Append "mon_des", adChar, 4, adFldIsNullable
RsTipCamb.Fields.Append "factor", adChar, 1, adFldIsNullable
RsTipCamb.Fields.Append "compra", adCurrency, 18, adFldIsNullable
RsTipCamb.Fields.Append "contable", adCurrency, 18, adFldIsNullable
RsTipCamb.Fields.Append "fec_crea", adDate, 10, adFldIsNullable
RsTipCamb.Open

Set GrdTipCamb.DataSource = RsTipCamb
If rs.State = 1 Then rs.Close

End Sub

Private Sub Form_Deactivate()
SQL$ = "SELECT compra,contable,factor from tipo_cambio where fec_crea " _
& " BETWEEN '" & Format(Date, FormatFecha) + FormatTimei & "' AND '" & Format(Date, FormatFecha) + FormatTimef & "' " _
& " AND cod_cia='" & wcia & "'"
        
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(SQL$, 64)
If rs.RecordCount > 0 Then
   mtipo_cambio = rs!compra
   mtc_contable = rs!contable
   findmon = True
Else
   MsgBox "No se Puede Cerrar,hasta que no Grabe un Tipo de Cambio del Día", vbInformation
   Me.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim wciamae As String
Me.Top = 0
Me.Left = 0
Me.Width = 8175
Me.Height = 4260
Lblfecha.Caption = Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm:ss AMPM")
windex = Val(wcia) - 1
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Cmbcia.ListIndex = windex
wciamae = Determina_Maestro("01006")
SQL$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where status=''"
SQL$ = SQL$ & wciamae
If (fAbrRst(rs, SQL$)) Then
   rs.MoveFirst
   Do Until rs.EOF
      lstmon.AddItem rs!FLAG1 & Space(1) & rs!descrip
      lstmon.ItemData(lstmon.NewIndex) = Trim(rs!cod_maestro2)
      rs.MoveNext
   Loop
End If
If rs.State = 1 Then rs.Close
Call rCargaDatos
End Sub
Private Sub Form_Unload(Cancel As Integer)
SQL$ = "SELECT compra,contable,factor from tipo_cambio where fec_crea " _
& " BETWEEN '" & Format(Date, FormatFecha) + FormatTimei & "' AND '" & Format(Date, FormatFecha) + FormatTimef & "' " _
& " AND cod_cia='" & wcia & "'"
        
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(SQL$, 64)
If rs.RecordCount > 0 Then
   mtipo_cambio = rs!compra
   mtc_contable = rs!contable
   findmon = True
Else
   MsgBox "No se Puede Cerrar,hasta que no Grabe un Tipo de Cambio del Día", vbInformation
   Cancel = True
End If
End Sub

Private Sub GrdTipCamb_AfterColEdit(ByVal ColIndex As Integer)
Dim mcambio As Currency
Select Case ColIndex
       Case Is = 3
            If GrdTipCamb.Columns(3) = "" Then GrdTipCamb.Columns(3) = 0
            mcambio = GrdTipCamb.Columns(3)
            RsTipCamb.MoveFirst
            Do While Not RsTipCamb.EOF
               RsTipCamb!compra = mcambio
               RsTipCamb.MoveNext
            Loop
       Case Is = 4
            If GrdTipCamb.Columns(4) = "" Then GrdTipCamb.Columns(4) = 0
            mcambio = GrdTipCamb.Columns(4)
            RsTipCamb.MoveFirst
            Do While Not RsTipCamb.EOF
               RsTipCamb!contable = mcambio
               RsTipCamb.MoveNext
            Loop
End Select
RsTipCamb.MoveFirst
GrdTipCamb.Row = 0
GrdTipCamb.Col = 3
GrdTipCamb.Refresh
End Sub

Private Sub GrdTipCamb_ButtonClick(ByVal ColIndex As Integer)
If Trim(COL7) <> "" Then
  If Format(GrdTipCamb.Columns(5), FormatFecha) = Format(Date, FormatFecha) Then
  GoTo salir
  Else
   Exit Sub
  End If
End If

salir:
Dim Y As Integer, xtop As Integer, xleft As Integer
Y = GrdTipCamb.Row
xtop = GrdTipCamb.Top + GrdTipCamb.RowTop(Y) + GrdTipCamb.RowHeight
Select Case GrdTipCamb.Col
Case 0, 1
xleft = GrdTipCamb.Left + GrdTipCamb.Columns(0).Left
       If Y < 12 Then
         lstmon.Top = xtop
        Else
         lstmon.Top = GrdTipCamb.Top + GrdTipCamb.RowTop(Y) - lstmon.Height
        End If
        If ColIndex = 0 Then lstmon.Left = xleft Else lstmon.Left = GrdTipCamb.Left + GrdTipCamb.Columns(1).Left
         
        lstmon.Visible = True
        lstmon.SetFocus

Case 2
      xleft = GrdTipCamb.Left + GrdTipCamb.Columns(2).Left
       With lstfactor
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = GrdTipCamb.Top + grdtipo.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub GrdTipCamb_KeyPress(KeyAscii As Integer)
Select Case GrdTipCamb.Col
        Case 0
         If Len(GrdTipCamb.Columns(2)) >= 1 And KeyAscii <> 8 Or KeyAscii = 39 Then
             KeyAscii = 0
         Else
             If KeyAscii = 47 Or KeyAscii = 42 Or KeyAscii = 8 Then Else: KeyAscii = 0
         End If
        Case 1
         If Len(GrdTipCamb.Columns(2)) >= 1 And KeyAscii <> 8 Or KeyAscii = 39 Then
             KeyAscii = 0
         Else
             If KeyAscii = 47 Or KeyAscii = 42 Or KeyAscii = 8 Then Else: KeyAscii = 0
         End If
        Case 2
         If Len(GrdTipCamb.Columns(2)) >= 1 And KeyAscii <> 8 Or KeyAscii = 39 Then
             KeyAscii = 0
         Else
             If KeyAscii = 47 Or KeyAscii = 42 Or KeyAscii = 8 Then Else: KeyAscii = 0
         End If
End Select
End Sub

Private Sub lstfactor_Click()
  GrdTipCamb.Columns(2).Text = lstfactor
  lstfactor.Visible = False
End Sub

Private Sub lstmon_Click()
If GrdTipCamb.Col = 0 Then GrdTipCamb.Columns(0) = Mid$(lstmon, 2, 3)
If GrdTipCamb.Col = 1 Then GrdTipCamb.Columns(1) = Mid$(lstmon, 2, 3)
lstmon.Visible = False
End Sub

Private Sub rCargaDatos()
Dim fecha As String

SQL$ = "SELECT " & FechaSys & " "
Set rs = New ADODB.Recordset
Set rs = cn.Execute(SQL$)
fecha = rs(0)
If rs.State = 1 Then rs.Close

SQL$ = "SELECT * from tipo_cambio where fec_crea " _
     & " BETWEEN '" & Format(Date, FormatFecha) + FormatTimei & "' AND '" & Format(Date, FormatFecha) + FormatTimef & "' " _
     & " AND cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, SQL$)) Then
   rs.MoveFirst
   Do While Not rs.EOF
      RsTipCamb.AddNew
      RsTipCamb!MON_ORI = rs!MON_ORI
      RsTipCamb!MON_DES = rs!MON_DES
      RsTipCamb!factor = rs!factor
      RsTipCamb!compra = rs!compra
      RsTipCamb!contable = rs!contable
      RsTipCamb!fec_crea = rs!fec_crea
      RsTipCamb.Update
      rs.MoveNext
   Loop
Else
   RsTipCamb.AddNew
   RsTipCamb!MON_ORI = "S/."
   RsTipCamb!MON_DES = "US$"
   RsTipCamb!factor = "/"
   RsTipCamb!compra = 0
   RsTipCamb!contable = 0
   RsTipCamb!fec_crea = fecha
   
   RsTipCamb.AddNew
   RsTipCamb!MON_ORI = "US$"
   RsTipCamb!MON_DES = "S/."
   RsTipCamb!factor = "*"
   RsTipCamb!compra = 0
   RsTipCamb!contable = 0
   RsTipCamb!fec_crea = fecha
End If
Set GrdTipCamb.DataSource = RsTipCamb
If rs.State = 1 Then rs.Close

GrdTipCamb.Columns(0).Width = 810
GrdTipCamb.Columns(1).Width = 810
GrdTipCamb.Columns(2).Width = 689
GrdTipCamb.Columns(3).Width = 1305
GrdTipCamb.Columns(4).Width = 1275
GrdTipCamb.Columns(5).Width = 2415

GrdTipCamb.Columns(0).Alignment = dbgCenter
GrdTipCamb.Columns(1).Alignment = dbgCenter
GrdTipCamb.Columns(2).Alignment = dbgCenter
GrdTipCamb.Columns(3).Alignment = dbgRight
GrdTipCamb.Columns(4).Alignment = dbgRight
GrdTipCamb.Columns(5).Alignment = dbgCenter

GrdTipCamb.Columns(3).NumberFormat = "####.000"
GrdTipCamb.Columns(4).NumberFormat = "####.000"



GrdTipCamb.Columns(0).Button = True
GrdTipCamb.Columns(1).Button = True
GrdTipCamb.Columns(2).Button = True

End Sub


