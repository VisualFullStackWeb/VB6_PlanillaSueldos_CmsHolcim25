VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmtareo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tareo Diario"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "Frmtareo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbccosto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1675
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Txtfecha 
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frametareo 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   8655
      Begin VB.ListBox Lstunidad 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "Frmtareo.frx":030A
         Left            =   6720
         List            =   "Frmtareo.frx":0317
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ListBox LstConceptos 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
      End
      Begin MSDataGridLib.DataGrid Dgdtareo 
         Height          =   4335
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "concepto"
            Caption         =   "Conceptos"
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
            DataField       =   "tiempo"
            Caption         =   "Tiempo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "unidad"
            Caption         =   "Unidad de Tiempo"
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
            DataField       =   "codigo"
            Caption         =   "Codigo"
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
            DataField       =   "Defecto"
            Caption         =   "Defecto"
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
         BeginProperty Column05 
            DataField       =   "Valor"
            Caption         =   "Valor"
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
               Button          =   -1  'True
               ColumnWidth     =   5174.929
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               Button          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   8655
      Begin VB.TextBox Txtcodobra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Txtcodtrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Lblobra 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Lblnombre 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6960
         TabIndex        =   16
         Top             =   120
         Width           =   1695
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
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Lbltipo 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "C. de Costo"
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "Frmtareo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstareo As New Recordset
Dim wciamae As String
Dim VArea As String
Dim rs2 As ADODB.Recordset
Private Sub Cmbccosto_Click()
VArea = fc_CodigoComboBox(CmbCcosto, 3)
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Crea_Rs
LstConceptos.Clear
Call fc_Descrip_Maestros2("01044", "", CmbCcosto)
End Sub

Private Sub Dgdtareo_AfterColEdit(ByVal ColIndex As Integer)
Dim mdeci As Integer
Select Case ColIndex
       Case Is = 1
            If Not IsNumeric(Dgdtareo.Columns(1)) Then
               Dgdtareo.Columns(1) = "0.00"
            Else
               'If Dgdtareo.Columns(1) <> Dgdtareo.Columns(5) And Dgdtareo.Columns(5) <> "0" Then
                '  MsgBox "Tiempo no puede ser Diferente de " & Str(Val(Dgdtareo.Columns(5))), vbInformation, "Tareo"
                  'Dgdtareo.Columns(1) = Dgdtareo.Columns(5)
               'End If
        
               If Dgdtareo.Columns(1) < 0 Then
                  MsgBox "Tiempo no puede ser Negativo", vbInformation, "Tareo"
                  Dgdtareo.Columns(1) = 0
               End If
               If Left(Dgdtareo.Columns(2), 2) = "DI" Or Left(Dgdtareo.Columns(2), 2) = "MI" Then
                  Dgdtareo.Columns(1) = Int(Dgdtareo.Columns(1))
               End If
               mdeci = Mid(Dgdtareo.Columns(1), InStr(1, Dgdtareo.Columns(1), ".", 1) + 1, 2)
               If mdeci > 60 Then
                  MsgBox "Decimales no pueden Exceder a 60", vbInformation, "Tareo"
                  Dgdtareo.Columns(1) = Int(Dgdtareo.Columns(1))
               End If
               If Left(Dgdtareo.Columns(2), 2) = "HO" And Dgdtareo.Columns(1) > 24 Then
                  MsgBox "Horas no pueden Exceder a 24", vbInformation, "Tareo"
                  Dgdtareo.Columns(1) = "0.00"
               End If
               If Left(Dgdtareo.Columns(2), 2) = "DI" And Dgdtareo.Columns(1) > 1 Then
                  MsgBox "Solo se puede registrar un dia", vbInformation, "Tareo"
                  Dgdtareo.Columns(1) = "1.00"
               End If
               If Left(Dgdtareo.Columns(2), 2) = "DI" And Dgdtareo.Columns(1) > 1 Then
                  MsgBox "Solo se puede registrar un dia", vbInformation, "Tareo"
                  Dgdtareo.Columns(1) = "1.00"
               End If
            End If
       Case Is = 2
            If Dgdtareo.Columns(2) <> Dgdtareo.Columns(4) And Trim(Dgdtareo.Columns(4)) <> "" Then
               MsgBox "Unidad no puede ser diferente de " & Dgdtareo.Columns(4), vbInformation, "Tareo"
               Dgdtareo.Columns(2) = Dgdtareo.Columns(4)
            End If
End Select
End Sub

Private Sub Dgdtareo_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Dgdtareo.Col = 0 Or Dgdtareo.Col = 2 Then
        KeyAscii = 0
        Cancel = True
        Dgdtareo_ButtonClick (ColIndex)
End If
End Sub

Private Sub Dgdtareo_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer
Y = Dgdtareo.Row
xtop = Dgdtareo.Top + Dgdtareo.RowTop(Y) + Dgdtareo.RowHeight
Select Case ColIndex
Case 0:
       xleft = Dgdtareo.Left + Dgdtareo.Columns(0).Left
       With LstConceptos
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = Dgdtareo.Top + Dgdtareo.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
Case 2:
       xleft = Dgdtareo.Left + Dgdtareo.Columns(2).Left
       With Lstunidad
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = Dgdtareo.Top + Dgdtareo.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub Dgdtareo_LostFocus()
Dim mdeci As Integer
If Trim(Dgdtareo.Columns(1)) <> "" Then
   If Left(Dgdtareo.Columns(2), 2) = "DI" Or Left(Dgdtareo.Columns(2), 2) = "MI" Then
      Dgdtareo.Columns(1) = Int(Dgdtareo.Columns(1))
   End If
End If
If Dgdtareo.Columns(1) <> "" Then
   mdeci = Mid(Dgdtareo.Columns(1), InStr(1, Dgdtareo.Columns(1), ".", 1) + 1, 2)
   If mdeci > 60 Then
      MsgBox "Decimales no pueden Exceder a 60", vbInformation, "Tareo"
      Dgdtareo.Columns(1) = Int(Dgdtareo.Columns(1))
   End If
End If
If Dgdtareo.Columns(2) <> "" And Dgdtareo.Columns(1) <> "" Then
   If Left(Dgdtareo.Columns(2), 2) = "HO" And Dgdtareo.Columns(1) > 24 Then
      MsgBox "Horas no pueden Exceder a 24", vbInformation, "Tareo"
      Dgdtareo.Columns(1) = "0.00"
   End If
End If
If Dgdtareo.Columns(2) <> "" And Dgdtareo.Columns(1) <> "" Then
   If Left(Dgdtareo.Columns(2), 2) = "DI" And Dgdtareo.Columns(1) > 1 Then
      MsgBox "Solo se puede registrar un dia", vbInformation, "Tareo"
      Dgdtareo.Columns(1) = "1.00"
   End If
End If
End Sub

Private Sub Dgdtareo_OnAddNew()
rstareo.AddNew
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 8775
Me.Height = 7020
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
LblFecha.Caption = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Frmgrdtareo.Procesa_Tareo
End Sub

Private Sub LstConceptos_Click()
Dim MREC As Integer
Dim munidad As String

MREC = rstareo.AbsolutePosition
If LstConceptos.ListIndex > -1 Then
   If rstareo.RecordCount > 0 Then rstareo.MoveFirst
   munidad = ""
   Do While Not rstareo.EOF
      If Format(Right(LstConceptos.Text, 2), "00") = Trim(rstareo!codigo) Then
         MsgBox "Concepto Seleccionado ya se encuentra registrado", vbInformation, "Tareo"
         rstareo.AbsolutePosition = MREC
         Dgdtareo.Col = 0
         Dgdtareo.SetFocus
         LstConceptos.ZOrder 1
         LstConceptos.Visible = False
         Exit Sub
      End If
      If Left(rstareo!unidad, 2) = "DI" Then
         MsgBox "Ya se registro informacion por 1 Dia", vbInformation, "Tareo"
         rstareo.AbsolutePosition = MREC
         Dgdtareo.Col = 0
         Dgdtareo.SetFocus
         LstConceptos.ZOrder 1
         LstConceptos.Visible = False
         Exit Sub
      End If
      If Trim(rstareo!unidad) <> "" Then munidad = rstareo!unidad
      rstareo.MoveNext
   Loop
   rstareo.AbsolutePosition = MREC
    Dim m, P As Integer
    m = Len(LstConceptos.Text) - 2
    rstareo!concepto = Trim(Left(LstConceptos.Text, m))
    rstareo!codigo = Format(Right(RTrim(LstConceptos.Text), 2), "00")
   
    wciamae = Determina_Maestro("01077")
    Sql$ = "Select * from maestros_2 where cod_maestro2='" & Format(Right(LstConceptos.Text, 2), "00") & "' and status<>'*'"
    Sql$ = Sql$ & wciamae
    If (fAbrRst(rs, Sql$)) Then
       If rs!flag3 = "DI" Then rstareo!unidad = "DIAS"
       If rs!flag3 = "HO" Then rstareo!unidad = "HORAS"
       If rs!flag3 = "MI" Then rstareo!unidad = "MINUTOS"
       If rs!flag3 = "DI" And rs!flag5 = "S" Then rstareo!defecto = "DIAS"
       If rs!flag3 = "HO" And rs!flag5 = "S" Then rstareo!defecto = "HORAS"
       If rs!flag3 = "MI" And rs!flag5 = "S" Then rstareo!defecto = "MINUTOS"
       rstareo!VALOR = 0
       If Not IsNull(rs!flag6) Then
          If Val(rs!flag6) > 0 Then rstareo!VALOR = Val(rs!flag6): rstareo!tiempo = Val(rs!flag6)
       End If
       If munidad <> "" And rs!flag3 = "DI" Then
          If rs!flag5 = "S" Then
             MsgBox "No puede Registrar concepto por un dia, excederia las 24 horas", vbInformation, "Tareo"
             rstareo!unidad = ""
             rstareo!concepto = ""
             rstareo!tiempo = "0.00"
             rstareo!codigo = ""
             rstareo!VALOR = "0.00"
             rstareo!defecto = ""
          Else
             rstareo!unidad = ""
          End If
       End If
    End If
    Dgdtareo.Col = 0
    Dgdtareo.SetFocus
    LstConceptos.ZOrder 1
    LstConceptos.Visible = False
End If
End Sub

Private Sub LstConceptos_LostFocus()
LstConceptos.Visible = False
End Sub

Private Sub Lstunidad_Click()
Dim MREC As Integer
MREC = rstareo.AbsolutePosition
If Lstunidad.ListIndex > -1 Then
   If rstareo.RecordCount > 0 Then rstareo.MoveFirst
   munidad = ""
   Do While Not rstareo.EOF
      If Left(Lstunidad.Text, 2) = "DI" And Trim(rstareo!unidad <> "") And MREC <> rstareo.AbsolutePosition Then
         MsgBox "No puede Registrar concepto por un dia, excederia las 24 horas", vbInformation, "Tareo"
         rstareo.AbsolutePosition = MREC
         rstareo!unidad = ""
         rstareo!tiempo = "0.00"
         Dgdtareo.Col = 2
         Dgdtareo.SetFocus
         Lstunidad.ZOrder 1
         Lstunidad.Visible = False
         Exit Sub
      End If
      rstareo.MoveNext
   Loop
   rstareo.AbsolutePosition = MREC

   If Trim(Lstunidad.Text) <> Trim(Dgdtareo.Columns(4)) And Trim(Dgdtareo.Columns(4)) <> "" Then
      MsgBox "Unidad no puede ser diferente de " & Dgdtareo.Columns(4), vbInformation, "Tareo"
   Else
      Dgdtareo.Columns(2) = Trim(Lstunidad.Text)
   End If
   If Left(Lstunidad.Text, 2) = "DI" Or Left(Lstunidad.Text, 2) = "MI" Then
      If Dgdtareo.Columns(1) <> "" Then Dgdtareo.Columns(1) = Int(Dgdtareo.Columns(1))
   End If
   Dgdtareo.Col = 2
   Dgdtareo.SetFocus
   Lstunidad.ZOrder 1
   Lstunidad.Visible = False
End If
End Sub

Private Sub Lstunidad_LostFocus()
Lstunidad.Visible = False
End Sub

Private Sub Txtcodobra_GotFocus()
wbus = "OB"
NameForm = "Frmtareo"
End Sub

Private Sub Txtcodobra_KeyPress(KeyAscii As Integer)
Txtcodobra.Text = Txtcodobra.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Txtcodobra.Text = Format(Txtcodobra.Text, "00000000"):  txtfecha.SetFocus
End Sub

Private Sub Txtcodobra_LostFocus()
wbus = ""
If Txtcodobra.Text <> "" Then
   Sql$ = "select cod_obra,descrip,status from plaobras where cod_cia='" & wcia & "' and cod_obra='" & Txtcodobra.Text & "' order by status"
   If (fAbrRst(rs, Sql$)) Then
      If rs!status = "*" Then
         MsgBox "Obra Eliminada", vbInformation, "Registro de Personal"
         Lblobra.Caption = ""
         Txtcodobra.SetFocus
      Else
         Lblobra.Caption = Trim(rs!DESCRIP)
         Call Verifica_Tareo(Trim(TxtCodTrab.Text), Trim(Txtcodobra.Text), txtfecha.Text, VArea)
      End If
   Else
     MsgBox "Codigo de Obra no Registrada", vbInformation, "Registro de Personal"
     Lblobra.Caption = ""
     'Txtcodobra.SetFocus
   End If
End If
End Sub

Private Sub TxtCodTrab_GotFocus()
wbus = "PL"
End Sub

Private Sub TxtCodTrab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtfecha.SetFocus
End Sub

Private Sub TxtCodTrab_LostFocus()
Dim mcod As String
Dim marea As String

LstConceptos.Clear
If Trim(TxtCodTrab.Text) <> "" Then
    Sql$ = nombre()
    Sql$ = Sql$ + "tipotrabajador,status,area from planillas where cia='" & wcia & "' AND placod='" & Trim(TxtCodTrab.Text) & "' order by status"
    marea = ""
    cn.CursorLocation = adUseClient
    Set rs = New ADODB.Recordset
    Set rs = cn.Execute(Sql$)
    If rs.RecordCount > 0 Then
       marea = rs!Area
       If rs!status = "*" Then
          MsgBox "Trabajador Eliminado", vbExclamation, "Codigo N° => " & TxtCodTrab.Text
          LblNombre.Caption = ""
          Lbltipo.Caption = ""
          TxtCodTrab.SetFocus
       Else
          wciamae = Determina_Maestro("01077")
          Sql$ = "Select cod_maestro2,descrip from maestros_2 a,plaverhoras b where a.status<>'*' and b.status<>'*' and b.cia='" & wcia & "' " _
               & " and tipo_trab='" & rs!TipoTrabajador & "' and b.codigo=a.cod_maestro2"
          Sql$ = Sql$ & wciamae
          cn.CursorLocation = adUseClient
          Set rs2 = New ADODB.Recordset
          Set rs2 = cn.Execute(Sql$, 64)
          If rs2.RecordCount > 0 Then rs2.MoveFirst
          Do Until rs2.EOF
             LstConceptos.AddItem rs2!DESCRIP & Space(100) & rs2!COD_MAESTRO2
             rs2.MoveNext
          Loop
          If rs2.State = 1 Then rs2.Close
          LblNombre.Caption = rs!nombre
          Lbltipo.Caption = rs!TipoTrabajador
          Call Verifica_Tareo(Trim(TxtCodTrab.Text), Trim(Txtcodobra.Text), txtfecha.Text, VArea)
       End If
    Else
       MsgBox "Codigo de Trabajador no Registrado", vbExclamation, "Codigo N° => " & TxtCodTrab.Text
       LblNombre.Caption = ""
       Lbltipo.Caption = ""
       TxtCodTrab.SetFocus
    End If
Else
   LblNombre.Caption = ""
   Lbltipo.Caption = ""
End If
wbus = ""
If Lbltipo.Caption = "05" Then
   Label5.Visible = True
   Txtcodobra.Visible = True
   Lblobra.Visible = True
Else
   Label5.Visible = False
   Txtcodobra.Visible = False
   Lblobra.Visible = False
End If

Call rUbiIndCmbBox(CmbCcosto, marea, "00")
If rs.State = 1 Then rs.Close
End Sub
Private Sub Crea_Rs()
    If rstareo.State = 1 Then rstareo.Close
    rstareo.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rstareo.Fields.Append "concepto", adChar, 300, adFldIsNullable
    rstareo.Fields.Append "unidad", adChar, 15, adFldIsNullable
    rstareo.Fields.Append "tiempo", adCurrency, 18, adFldIsNullable
    rstareo.Fields.Append "defecto", adChar, 15, adFldIsNullable
    rstareo.Fields.Append "valor", adCurrency, 18, adFldIsNullable
    rstareo.Open
    rstareo.AddNew
    Set Dgdtareo.DataSource = rstareo
End Sub
Public Sub Grabar_Tareo()
Dim mmonto As Currency
Dim mdeci As Integer
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

If rstareo.RecordCount <= 0 Then Exit Sub
If Trim(TxtCodTrab.Text) = "" Then MsgBox "Debe Ingresar Codigo de Trabajador", vbInformation, "Tareo": TxtCodTrab.SetFocus: Exit Sub
If Trim(Txtcodobra.Text) = "" And Lbltipo.Caption = "05" Then MsgBox "Debe Ingresar Codigo de Obra", vbInformation, "Tareo": Txtcodobra.SetFocus: Exit Sub
If Not IsDate(txtfecha.Text) Then MsgBox "Debe Ingresar Fecha", vbInformation, "Tareo": txtfecha.SetFocus: Exit Sub
If Trim(CmbCcosto.Text) = "" Then MsgBox "Debe Seleccionar C. de Costo", vbInformation, "Tareo": CmbCcosto.SetFocus: Exit Sub

rstareo.MoveFirst
Do While Not rstareo.EOF
   If Not IsNull(rstareo!tiempo) Then
   If Left(rstareo!unidad, 2) = "DI" Or Left(rstareo!unidad, 2) = "MI" Then
      rstareo!tiempo = Int(rstareo!tiempo)
   End If
   mdeci = Mid(rstareo!tiempo, InStr(1, rstareo!tiempo, ".", 1) + 1, 2)
   If mdeci > 60 And InStr(1, rstareo!tiempo, ".", 1) > 0 Then
      MsgBox "Decimales no pueden Exceder a 60", vbInformation, "Tareo"
      Exit Sub
   End If
   If rstareo!tiempo <> rstareo!VALOR And rstareo!VALOR <> 0 Then
      MsgBox "Tiempo no puede ser Diferente de " & Str(rstareo!defecto), vbInformation, "Tareo"
      rstareo!tiempo = rstareo!defecto
      Exit Sub
   End If
   If Left(rstareo!unidad, 2) = "HO" And rstareo!tiempo > 24 Then
      MsgBox "Horas no pueden Exceder a 24", vbInformation, "Tareo"
      rstareo!tiempo = "0.00"
      Exit Sub
   End If
   If Left(rstareo!unidad, 2) = "DI" And rstareo!tiempo > 1 Then
      MsgBox "Solo se puede registrar un dia", vbInformation, "Tareo"
      rstareo!tiempo = "1.00"
      Exit Sub
   End If
   End If
  rstareo.MoveNext
Loop

If Validar_Tareo() = False Then Exit Sub

If MsgBox("Desea Grabar Tareo ", vbYesNo + vbQuestion) = vbNo Then Exit Sub

Screen.MousePointer = vbHourglass
If rstareo.RecordCount > 0 Then rstareo.MoveFirst

cn.BeginTrans
NroTrans = 1

Sql$ = "update platareo set status='*' where cia='" & wcia & "' and codigotrab='" & TxtCodTrab.Text & "' and obra='" & Txtcodobra.Text & "' and fecha='" & Format(txtfecha.Text, FormatFecha) & "' and status<>'*'"
cn.Execute Sql$

Do While Not rstareo.EOF
   If IsNull(rstareo!tiempo) Then mmonto = 0 Else mmonto = rstareo!tiempo
   If CCur(mmonto) <> 0 Then
      Sql$ = "set dateformat " & Coneccion.FormatFechaSql & " "
      Sql$ = Sql$ & "Insert into platareo values('" & wcia & _
      "','" & Format(txtfecha.Text, FormatFecha) & "','" & _
      Txtcodobra.Text & "','" & Trim(VArea) & "','" & _
      Trim(TxtCodTrab.Text) & "','" & rstareo!codigo & _
      "'," & CCur(mmonto) & ",'" & Left(rstareo!unidad, 2) & _
      "','','" & wuser & "'," & FechaSys & ",''," & FechaSys & ")"
      cn.Execute Sql$
   End If
   rstareo.MoveNext
Loop

Txtcodobra.Text = ""
Lblobra.Caption = ""
If rstareo.RecordCount > 0 Then
   rstareo.MoveFirst
   Do While Not rstareo.EOF
      rstareo.Delete
      rstareo.MoveNext
   Loop
   rstareo.AddNew
End If

cn.CommitTrans
Screen.MousePointer = vbDefault
MsgBox "Tareo Grabado Exitosamente", vbInformation, "Planillas"
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
Screen.MousePointer = vbDefault
MsgBox ERR.Description, vbInformation, "Planillas"


End Sub

Private Sub Txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmbCcosto.SetFocus
End Sub

Private Sub Txtfecha_LostFocus()
If txtfecha.Text <> "__/__/____" And Not IsDate(txtfecha.Text) Then
   MsgBox "Ingrese Correctamente la Fecha", vbInformation, "Tareo"
   txtfecha.SetFocus
End If
Call Verifica_Tareo(Trim(TxtCodTrab.Text), Trim(Txtcodobra.Text), txtfecha.Text, VArea)
End Sub
Public Sub Nuevo_Tareo()
TxtCodTrab.Text = ""
LblNombre.Caption = ""
Txtcodobra.Text = ""
Lblobra.Caption = ""
CmbCcosto.ListIndex = -1
If rstareo.RecordCount > 0 Then
   rstareo.MoveFirst
   Do While Not rstareo.EOF
      rstareo.Delete
      rstareo.MoveNext
   Loop
   rstareo.AddNew
End If
End Sub
Public Sub Carga_Tareo(codtrab As String, obra As String, fecha As String, ccosto As String)
TxtCodTrab.Text = codtrab
Txtcodobra.Text = obra
txtfecha.Text = Format(fecha, "dd/mm/yyyy")
TxtCodTrab_LostFocus
Txtcodobra_LostFocus
Call rUbiIndCmbBox(CmbCcosto, ccosto, "000")
If rstareo.RecordCount > 0 Then
   rstareo.MoveFirst
   Do While Not rstareo.EOF
      rstareo.Delete
      rstareo.MoveNext
   Loop
End If
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "select * from platareo where cia='" & wcia & "' and codigotrab='" & codtrab & "' and obra='" & obra & "' " _
     & "and fecha='" & Format(fecha, FormatFecha) & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   Do While Not rs.EOF
       rstareo.AddNew
       rstareo!codigo = rs!concepto
       If rs!motivo = "DI" Then rstareo!unidad = "DIAS"
       If rs!motivo = "HO" Then rstareo!unidad = "HORAS"
       If rs!motivo = "MI" Then rstareo!unidad = "MINUTOS"
       rstareo!tiempo = rs!tiempo
       rs.MoveNext
   Loop
End If
If rs.State = 1 Then rs.Close
If rstareo.RecordCount > 0 Then
   rstareo.MoveFirst
   Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='077' and status<>'*' "
   If (fAbrRst(rs, Sql$)) Then
      If rs!General = "S" Then
         wciamae = " and right(ciamaestro,3)= '077' ORDER BY cod_maestro2"
      Else
         wciamae = " and ciamaestro= '" & wcia + "077" & "' ORDER BY cod_maestro2"
      End If
   End If
   If rs.State = 1 Then rs.Close
   Do While Not rstareo.EOF
      Sql$ = "SELECT * From maestros_2 where cod_maestro2='" & rstareo!codigo & "'"
      Sql$ = Sql$ & wciamae
      Set rs = cn.Execute(Sql$)
      If rs.RecordCount > 0 Then
         rstareo!concepto = rs!DESCRIP
         If rs!flag3 = "DI" Then rstareo!unidad = "DIAS"
         If rs!flag3 = "HO" Then rstareo!unidad = "HORAS"
         If rs!flag3 = "MI" Then rstareo!unidad = "MINUTOS"
         If rs!flag3 = "DI" And rs!flag5 = "S" Then rstareo!defecto = "DIAS"
         If rs!flag3 = "HO" And rs!flag5 = "S" Then rstareo!defecto = "HORAS"
         If rs!flag3 = "MI" And rs!flag5 = "S" Then rstareo!defecto = "MINUTOS"
         rstareo!VALOR = "0.00"
         If Not IsNull(rs!flag6) Then
             If Val(rs!flag6) > 0 Then rstareo!VALOR = Val(rs!flag6): rstareo!tiempo = Val(rs!flag6)
          End If
      End If
      rstareo.MoveNext
   Loop
End If
If rstareo.RecordCount > 0 Then rstareo.MoveFirst
End Sub
Public Sub Verifica_Tareo(codtrab As String, obra As String, fecha As String, ccosto As String)
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "select * from platareo where cia='" & wcia & "' and codigotrab='" & Trim(codtrab) & "' and obra='" & obra & "' " _
     & "and fecha='" & Format(fecha, FormatFecha) & "' and status<>'*'"
     
If Not IsDate(fecha) Then Exit Sub
If (fAbrRst(rs, Sql$)) Then
   If rstareo.RecordCount > 0 Then
      rstareo.MoveFirst
      Do While Not rstareo.EOF
         rstareo.Delete
         rstareo.MoveNext
      Loop
   End If

   rs.MoveFirst
   Do While Not rs.EOF
       rstareo.AddNew
       rstareo!codigo = rs!concepto
       If rs!motivo = "DI" Then rstareo!unidad = "DIAS"
       If rs!motivo = "HO" Then rstareo!unidad = "HORAS"
       If rs!motivo = "MI" Then rstareo!unidad = "MINUTOS"
       rstareo!tiempo = rs!tiempo
       rs.MoveNext
   Loop
End If
If rs.State = 1 Then rs.Close
If rstareo.RecordCount > 0 Then
   rstareo.MoveFirst
   Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='077' and status<>'*' "
   If (fAbrRst(rs, Sql$)) Then
      If rs!General = "S" Then
         wciamae = " and right(ciamaestro,3)= '077' ORDER BY cod_maestro2"
      Else
         wciamae = " and ciamaestro= '" & wcia + "077" & "' ORDER BY cod_maestro2"
      End If
   End If
   If rs.State = 1 Then rs.Close
   Do While Not rstareo.EOF
      Sql$ = "SELECT * From maestros_2 where cod_maestro2='" & rstareo!codigo & "'"
      Sql$ = Sql$ & wciamae
      Set rs = cn.Execute(Sql$)
      If rs.RecordCount > 0 Then
         rstareo!concepto = rs!DESCRIP
         If rs!flag3 = "DI" Then rstareo!unidad = "DIAS"
         If rs!flag3 = "HO" Then rstareo!unidad = "HORAS"
         If rs!flag3 = "MI" Then rstareo!unidad = "MINUTOS"
         If rs!flag3 = "DI" And rs!flag5 = "S" Then rstareo!defecto = "DIAS"
         If rs!flag3 = "HO" And rs!flag5 = "S" Then rstareo!defecto = "HORAS"
         If rs!flag3 = "MI" And rs!flag5 = "S" Then rstareo!defecto = "MINUTOS"
         rstareo!VALOR = "0.00"
         If Not IsNull(rs!flag6) Then
             If Val(rs!flag6) > 0 Then rstareo!VALOR = Val(rs!flag6): rstareo!tiempo = Val(rs!flag6)
          End If
      End If
      rstareo.MoveNext
   Loop
End If
If rstareo.RecordCount > 0 Then rstareo.MoveFirst
End Sub
Private Function Validar_Tareo() As Boolean
Dim mhor As Currency
Validar_Tareo = True
If rstareo.RecordCount > 0 Then rstareo.MoveFirst
mhor = 0
Do While Not rstareo.EOF
   Select Case Left(rstareo!unidad, 2)
          Case Is = "DI"
               mhor = mhor + (rstareo!tiempo * 8 * 60)
          Case Is = "HO"
               mhor = mhor + Int(rstareo!tiempo) * 60 + ((rstareo!tiempo - Int(rstareo!tiempo)) * 100)
          Case Is = "MI"
               mhor = mhor + rstareo!tiempo
   End Select
   rstareo.MoveNext
Loop
mhor = Int(mhor / 60) + ((mhor Mod 60) / 100)
If mhor > 24 Then
   MsgBox "Registro de Tareo Contiene mas de 24 Horas", vbCritical, "Tareo"
   Validar_Tareo = False
End If
End Function
Public Sub Elimina_Tareo()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
Mgrab = MsgBox("Seguro de Eliminar Tareo", vbYesNo + vbQuestion, "Tareo Diario")
If Mgrab <> 6 Then Exit Sub
Screen.MousePointer = vbHourglass
cn.BeginTrans
NroTrans = 1
Sql$ = "update platareo set status='*' where cia='" & wcia & "' and codigotrab='" & TxtCodTrab.Text & "' and obra='" & Txtcodobra.Text & "' and fecha='" & Format(txtfecha.Text, FormatFecha) & "' and status<>'*'"
cn.Execute Sql$
cn.CommitTrans
MsgBox "Eliminación Satisfactoria", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault

End Sub
