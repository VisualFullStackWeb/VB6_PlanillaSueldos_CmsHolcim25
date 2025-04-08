VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmdesccta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos de Boletas"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   7710
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   7695
      Begin VB.TextBox txtsemana 
         Appearance      =   0  'Flat
         Height          =   380
         Left            =   3000
         TabIndex        =   12
         Top             =   285
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtfecha 
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   330
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   72810497
         CurrentDate     =   39085
      End
      Begin MSComCtl2.DTPicker DTPFIN 
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   72810497
         CurrentDate     =   39077
      End
      Begin MSComCtl2.DTPicker DTPIN 
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483624
         Format          =   72810497
         CurrentDate     =   39077
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   285
         Width           =   195
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
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
         Left            =   5760
         TabIndex        =   19
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
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
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inicio :"
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
         Left            =   2280
         TabIndex        =   17
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
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
         TabIndex        =   16
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtid 
         Height          =   285
         Left            =   5880
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtcod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtcuota 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbltot 
         Caption         =   "S\.000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Semanal : "
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
         Left            =   2280
         TabIndex        =   5
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblnom 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
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
         TabIndex        =   2
         Top             =   280
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dgdes 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4260
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483624
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID_DOC"
         Caption         =   "CODIGODOC"
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
         DataField       =   "YEAR"
         Caption         =   "AÑO"
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
         DataField       =   "SEMANA"
         Caption         =   "SEMANA"
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
         DataField       =   "CANTIDAD"
         Caption         =   "MONTO"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmdesccta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rtempo As New ADODB.Recordset
Public proc As Integer

'1   frmboleta
Sub Crea_Rs()
    With rtempo
         '.Fields.Append "codigo", adInteger, 4, adFldKeyColumn
         .Fields.Append "cantidad", adDouble, 2, adFldIsNullable
         .Fields.Append "semana", adInteger, 2
         .Fields.Append "year", adInteger, 4, adFldIsNullable
          
         rtempo.Open
         Set dgdes.DataSource = rtempo
    End With


End Sub

Public Sub guardar1()
On Error GoTo CORRIGE
Dim cadfecha As String
Dim con As String

If Val(Txtsemana) = 0 Then
   MsgBox "Ingrese La Semana", vbInformation, Me.Caption
   Exit Sub
End If

If Trim(txtcod) = "" Then
   MsgBox "Ingrese El Codigo del Trabajador", vbInformation, Me.Caption
   Exit Sub
End If

cadfecha = Str(Format(DTPFIN, "YYYY")) & "-" & Trim(Str(Format(Txtsemana, "00")))
       con = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " " & _
       "insert into pladescta values(" & Trim(lblnom.Tag) & "," & _
       Trim(txtcuota.Text) & ",'" & Trim(cadfecha) & "','','" & Format(dtfecha, "dd/mm/yyyy") & "')"

       cn.Execute con
   MsgBox "Datos Guardados correctamente", vbExclamation, Me.Caption
   If proc = 1 Then
      Call llena_grid
      Unload Me
   End If
   
   Call LIMPIA
Exit Sub
CORRIGE:
MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Sub
Private Sub Form_Load()
  proc = 0
  Call Crea_Rs
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
If rtempo.State Then rtempo.Close
Set rtempo = Nothing
End Sub

Public Sub Txtcod_KeyPress(KeyAscii As Integer)
Dim con As String
Dim Rt As New ADODB.Recordset

        
        If KeyAscii = 13 Then
        
con = " SELECT cta.*,rtrim(pla.ap_pat) + ' ' + LTRIM(rtrim(pla.ap_mat))+ ' '+ LTRIM(rtrim(pla.nom_1))+ " & _
"' '+ LTRIM(rtrim(pla.nom_2)) as nombre FROM PLACTACTE cta,planillas pla where " & _
"CTA.placod='" & Trim(txtcod) & "' AND cta.placod=pla.placod AND CTA.STATUS<>'*'"

           Rt.Open con, cn, adOpenStatic, adLockReadOnly
           If Rt.RecordCount > 0 Then
              
              lblnom.Tag = Rt("CODAUXINTERNO")
              lblnom = UCase(Rt("nombre"))
              lbltot.Caption = Rt("IMPORTE")
           Else
              Exit Sub
           End If
           Rt.Close
           
           'CON = "select * from  pladescta where CODAUXINTERNO=" & lblnom.Tag
           'RS.Open CON, cn, adOpenStatic, adLockBatchOptimistic
        
          ' If rs.RecordCount > 0 Then
           'Do While Not RS.EOF
            '     rtempo.AddNew
             '    rtempo("fecha") = RS("fecha")
              '   rtempo("cantidad") = RS("monto")
               '  RS.MoveNext
           'Loop
          ' Else
           '      rtempo.AddNew
            '     rtempo("FECHA") = ""
             '    rtempo("CANTIDAD") = 0
           'End If
           'RS.Close

        End If
End Sub

Private Sub txtcuota_KeyPress(KeyAscii As Integer)
Dim OP As Double
        If KeyAscii = 13 Then
        Exit Sub
           If Val(lbltot) = 0 Or Val(txtcuota) = 0 Then Exit Sub
           OP = Val(lbltot.Caption) / Val(txtcuota)
           If OP > 0 Then CALCULA_SEMANAS (OP)
        End If
End Sub

Sub CALCULA_SEMANAS(semana As Integer)
Dim i As Integer
Dim SEMFIN As Integer, NRO As Integer, Inicio As Integer
Dim con As String
Dim FLAG As Boolean
Dim RY As New ADODB.Recordset
Dim FINSEM As Integer
On Error GoTo CORRIGE

FLAG = False
If Val(Txtsemana.Text) = 0 Then
   MsgBox "Seleccione La Semana a Comenzar", vbInformation, Me.Caption
   Exit Sub
End If

FINSEM = Val(Txtsemana) + Val(semana)
i = 0
Inicio = Val(Txtsemana)
   Do While Not FLAG
     con = "SELECT * FROM PLASEMANAS WHERE CIA='" & wcia & "' AND " & _
     Format(DTPFIN.Value, "YYYY") & "=ANO AND SEMANA='" & Format(i + Inicio, "00") & _
     "' and status<>'*'"
    RY.Open con, cn, adOpenStatic, adLockReadOnly
    '========================================
     rtempo.AddNew
         rtempo("year") = Format(DTPFIN, "YYYY")
         If RY.RecordCount > 0 Then
            rtempo("SEMANA") = RY("SEMANA")
            rtempo("YEAR") = Format(DTPIN, "YYYY")
         Else
            FINSEM = FINSEM - (i + Inicio - 1)
            rtempo("SEMANA") = 1
            rtempo("YEAR") = Format(DTPIN, "YYYY") + 1
            DTPFIN.Value = "01/01/" & rtempo("YEAR")
            i = 0
            Inicio = 0
         End If
'             Debug.Print Val(txtcuota)
    
            rtempo("CANTIDAD") = Val(txtcuota)
     i = i + 1
     If FINSEM = (Inicio + i) Then FLAG = True
     RY.Close
   Loop
   
   Set RY = Nothing
   '==========================================
Exit Sub
CORRIGE:
MsgBox "Error : " & Err.Description, vbCritical, Me.Caption

End Sub

Private Sub Txtsemana_Change()
Dim SQL As String
Dim RX As New ADODB.Recordset

On Error GoTo CORRIGE

If Trim(Txtsemana) = "" Then Exit Sub


SQL$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(dtfecha, "YYYY") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
Set RX = cn.Execute(SQL)

If RX.RecordCount > 0 Then
   DTPIN.Value = Format(RX("fechai"), "DD/MM/YYYY")
   DTPFIN.Value = Format(RX("fechaF"), "DD/MM/YYYY")
End If

If RX.State = adStateOpen Then RX.Close

Exit Sub
CORRIGE:
MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub UpDown1_DownClick()
If Trim(Txtsemana.Text) = "" Then Txtsemana.Text = "0"
If Txtsemana.Text > 0 Then Txtsemana = Txtsemana - 1
End Sub

Private Sub UpDown1_UpClick()
If Trim(Txtsemana.Text) = "" Then Txtsemana.Text = "0"
Txtsemana = Txtsemana + 1
End Sub

Sub GUARDAR()
    On Error GoTo CORRIGE
    Dim i As Integer
    Dim con As String
    Dim cadfecha As String
     
    rtempo.MoveFirst
    Do While Not rtempo.EOF
    
       cadfecha = Format(rtempo("year"), "YYYY") & "-" & Format(rtempo("semana"), "00")
       con = "insert into pladescta values(" & lblnom.Tag & "," & _
       rtempo("cantidad") & "," & cadfecha & "' ')"
     
       rtempo.MoveNext
    Loop
    Exit Sub
CORRIGE:
    MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Sub

Sub llena_grid()

     If Frmboleta.rsdesadic.RecordCount > 0 Then
        Frmboleta.rsdesadic.MoveFirst
        Do While Not Frmboleta.rsdesadic.EOF
           If Frmboleta.rsdesadic.Fields("codigo") = "07" Then
              Frmboleta.rsdesadic.Fields("monto") = Val(txtcuota.Text)
           End If
           Frmboleta.rsdesadic.MoveNext
        Loop

     End If


End Sub

Sub LIMPIA()
    txtcod.Text = ""
    txtcod.Tag = ""
    txtcuota.Text = ""
    lbltot.Caption = ""
    Txtsemana.Text = ""
    lblnom.Caption = ""
End Sub
