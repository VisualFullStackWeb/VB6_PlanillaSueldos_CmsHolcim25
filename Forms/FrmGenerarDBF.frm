VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGenerarDBF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Generar Archivo DBF «"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "FrmGenerarDBF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4185
   Begin VB.Data Datapase 
      Connect         =   "FoxPro 2.5;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   975
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1725
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   75
      TabIndex        =   0
      Top             =   -45
      Width           =   3990
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   39314
      End
      Begin VB.TextBox TxtAño 
         Height          =   285
         Left            =   2325
         TabIndex        =   2
         Top             =   2775
         Width           =   1515
      End
      Begin VB.ComboBox CmbMes 
         Height          =   315
         ItemData        =   "FrmGenerarDBF.frx":030A
         Left            =   2325
         List            =   "FrmGenerarDBF.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   3150
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar P1 
         Height          =   165
         Left            =   150
         TabIndex        =   3
         Top             =   700
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione fecha a procesar"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   285
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese Año Proceso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   2820
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Mes Proceso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   3210
         Width           =   2010
      End
   End
End
Attribute VB_Name = "FrmGenerarDBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s_CodEmpresa_Starsoft As String
Dim rs_GenerarDBF As ADODB.Recordset
Dim rs_GenerarDBF2 As ADODB.Recordset
Dim s_MesSeleccion As String
Dim i_mesSeleccion As Integer
'Dim Importe_DBF As String
Dim Importe_DBF As Double
Dim s_Tipo_Asiento As String 'Haber o Debe
Dim s_Voucher As String
Dim s_Lote As String
Sub Generar_Archivo_Dbf()

Dim mOrigen As String
Dim Mdestino As String

i_mesSeleccion = Month(DTPicker1.Value)
Call Captura_Mes_Seleccionado
Call Codigo_Empresa_Starsoft

Call Enviar_Archivo_Origen(s_CodEmpresa_Starsoft)

Datapase.DatabaseName = Path_CIA & s_CodEmpresa_Starsoft & "\"

Datapase.RecordSource = "IMPORTAR.DBF"
Datapase.Refresh

Call Recupera_Informacion_ImportacionDBF(wcia, Year(DTPicker1.Value), s_MesSeleccion)
Set rs_GenerarDBF = Crear_Plan_Contable.rs_PlanCont_Pub
Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing

With rs_GenerarDBF
Do While Not .EOF
    Importe_DBF = 0: s_Voucher = "": s_Lote = ""
    If rs_GenerarDBF!Pla_Haber <> 0 Then
        'Importe_DBF = Str(rs_GenerarDBF!Pla_Haber)
        Importe_DBF = rs_GenerarDBF!Pla_Haber
    End If
    If rs_GenerarDBF!Pla_DEBE <> 0 Then
        'Importe_DBF = Str(rs_GenerarDBF!Pla_DEBE)
        Importe_DBF = rs_GenerarDBF!Pla_DEBE
    End If
    If Importe_DBF <> 0 Then
        Select Case Mid(rs_GenerarDBF!Pla_Cgcod, 1, 1)
            Case "9":
                s_Tipo_Asiento = "1"
                s_Tipo_Asiento = IIf(Trim(rs_GenerarDBF!Pla_Haber) > 0, "2", "1")
            Case "6": s_Tipo_Asiento = "1"
            Case "7": s_Tipo_Asiento = "2"
'            Case "4": s_Tipo_Asiento = "2"
'            Case "1": s_Tipo_Asiento = "2"
            Case "4"
                If Trim(rs_GenerarDBF!Pla_Cgcod) = "41300000" Then
                    s_Tipo_Asiento = IIf(Trim(rs_GenerarDBF!PLA_TIPO) = "0", "2", "1")
                Else
                    s_Tipo_Asiento = IIf(Trim(rs_GenerarDBF!Pla_DEBE) > 0, "1", "2")
                End If
            Case "1": s_Tipo_Asiento = "2"
        End Select
        Select Case rs_GenerarDBF!Pla_Boleta
            Case "01"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "PEN": s_Voucher = "0009"
                Else
                    s_Lote = "P0N": s_Voucher = "0013"
                End If
            Case "02"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "PEV": s_Voucher = "0009"
                Else
                    s_Lote = "P0V": s_Voucher = "0013"
                End If
            Case "03"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "PEG": s_Voucher = "0009"
                Else
                    s_Lote = "P0G": s_Voucher = "0013"
                End If
            Case "04"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "PEL": s_Voucher = "0009"
                Else
                    s_Lote = "P0L": s_Voucher = "0013"
                End If
            Case "12"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "VAE": s_Voucher = "0002"
                Else
                    s_Lote = "VAO": s_Voucher = "0002"
                End If
            Case "13"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "GRE": s_Voucher = "0002"
                Else
                    s_Lote = "GRO": s_Voucher = "0002"
                End If
            Case "14"
                If rs_GenerarDBF!Pla_TipTrabajador = "1" Then
                    s_Lote = "TSE": s_Voucher = "0002"
                Else
                    s_Lote = "TSO": s_Voucher = "0002"
                End If
        End Select
        
        If s_Voucher = "" Then
            MsgBox s_Voucher & " UN VOUCHER NO TIENE NUMERO. SE CANCLA LA OPERACION"
            Exit Sub
            
        End If
        
        Datapase.Recordset.AddNew
        Datapase.Recordset!cgcia = "01"
        Datapase.Recordset!CGCOD = rs_GenerarDBF!Pla_Cgcod
        Datapase.Recordset!CGDIA = DTPicker1.Value
        Datapase.Recordset!cgper = s_MesSeleccion
        Datapase.Recordset!CGIMP = Importe_DBF
        Datapase.Recordset!cgimpd = 0
        Datapase.Recordset!CGMOV = s_Tipo_Asiento
        Datapase.Recordset!CGVOUCH = s_Voucher
        Datapase.Recordset!CGLOTE = s_Lote
        Datapase.Recordset!cgnum = 0
        Datapase.Recordset!cgmon = "S"
        Datapase.Recordset!diario = "06"
        Datapase.Recordset!CENCOS = rs_GenerarDBF!Pla_CC
        Datapase.Recordset.Update
        
    End If
    .MoveNext
Loop
End With
Datapase.Refresh
MsgBox "Se genero el Archivo ", vbInformation
Unload Me
End Sub
Sub Codigo_Empresa_Starsoft()
'    Call Recuperar_Codigo_Empresa_Starsoft(wcia)
'    Set rs_GenerarDBF2 = Reportes_Centrales.rs_RptCentrales_pub
'    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
'    s_CodEmpresa_Starsoft = rs_GenerarDBF2!ciastar
'    Set rs_GenerarDBF2 = Nothing
    Call Trae_Cia_StarSoft(wcia, Format(DTPicker1.Value, "yyyy"))
    Set rs_GenerarDBF2 = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    s_CodEmpresa_Starsoft = rs_GenerarDBF2!EMP_ID
    Set rs_GenerarDBF2 = Nothing
End Sub
Sub Captura_Mes_Seleccionado()
    Select Case i_mesSeleccion
        Case 1: s_MesSeleccion = "01"
        Case 2: s_MesSeleccion = "02"
        Case 3: s_MesSeleccion = "03"
        Case 4: s_MesSeleccion = "04"
        Case 5: s_MesSeleccion = "05"
        Case 6: s_MesSeleccion = "06"
        Case 7: s_MesSeleccion = "07"
        Case 8: s_MesSeleccion = "08"
        Case 9: s_MesSeleccion = "09"
        Case 10: s_MesSeleccion = "10"
        Case 11: s_MesSeleccion = "11"
        Case 12: s_MesSeleccion = "12"
    End Select
End Sub
Private Sub Form_Activate()
    Call Llena_Barra
End Sub
Sub Llena_Barra()
    Dim i_Contador As Integer
    P1.Min = 1: P1.Max = 10
    For i_Contador = 1 To 10: P1.Value = i_Contador: Next i_Contador
End Sub
Sub Enviar_Archivo_Origen(Codigo_Empresa As String)
    Dim mOrigen As String
    Dim Mdestino As String
    mOrigen = App.Path & "\importar.dbf"
    Mdestino = Path_CIA & Codigo_Empresa & "\importar.dbf"
    'MsgBox " " & mOrigen & " Archivo Origen", vbInformation
    'MsgBox " " & Mdestino & "Archivo Destino", vbInformation
    CopyFile mOrigen, Mdestino, 0
End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
Me.Top = 0
Me.Left = 0
End Sub


