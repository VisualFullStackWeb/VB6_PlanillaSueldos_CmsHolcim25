VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmExpotar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Exportar a Asientos «"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "FrmExpotar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11775
   Begin VB.TextBox txt_Year 
      Height          =   315
      Left            =   975
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      ItemData        =   "FrmExpotar.frx":030A
      Left            =   1920
      List            =   "FrmExpotar.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   2385
   End
   Begin MSComctlLib.ProgressBar pbD 
      Height          =   105
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   4365
      Left            =   135
      TabIndex        =   3
      Top             =   795
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   7699
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
   Begin MSComctlLib.ProgressBar pbC 
      Height          =   150
      Left            =   135
      TabIndex        =   4
      Top             =   465
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSForms.CommandButton btn_Load 
      Height          =   375
      Left            =   4485
      TabIndex        =   13
      Top             =   30
      Width           =   2145
      Caption         =   "     Iniciar Exportación"
      PicturePosition =   327683
      Size            =   "3784;661"
      Picture         =   "FrmExpotar.frx":039A
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.SpinButton sp_Year 
      Height          =   315
      Left            =   1590
      TabIndex        =   12
      Top             =   60
      Width           =   255
      Size            =   "450;556"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   60
      Width           =   645
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   5220
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   2145
      TabIndex        =   9
      Top             =   5220
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Index           =   3
      Left            =   7860
      TabIndex        =   8
      Top             =   5235
      Width           =   420
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   8400
      TabIndex        =   7
      Top             =   5220
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   5
      Left            =   10410
      TabIndex        =   6
      Top             =   5220
      Width           =   1245
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Index           =   6
      Left            =   9810
      TabIndex        =   5
      Top             =   5220
      Width           =   480
   End
End
Attribute VB_Name = "FrmExpotar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim Cadena As String
Dim mMonth As String
Dim Cn_dbf As ADODB.Connection

Private Sub btn_Load_Click()
If MsgBox("Desea Exportar Movimientos ?", vbQuestion + vbYesNo, "Sistema") = vbYes Then
    Call Insert
End If
End Sub

Private Sub Cbo_Click()
    mMonth = Format(Cbo.ListIndex + 1, "00")
    Call Load_Info
End Sub

Private Sub Form_Load()
    Call Init_Form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Cn_dbf Is Nothing Then
        If Cn_dbf.State = adStateOpen Then Cn_dbf.Close
        Set Cn_dbf = Nothing
    End If
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then Ors.Close
        Set rs = Nothing
    End If
End Sub

Private Sub sp_Year_SpinDown()
    Call Down(txt_Year)
End Sub

Private Sub sp_Year_SpinUp()
    Call Up(txt_Year)
End Sub

Private Sub txt_Year_KeyPress(KeyAscii As Integer)
    Call NumberOnly(KeyAscii)
End Sub

Private Sub Load_Info()
    Cadena = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" & Path_CIA & Cia_StarSoft(wcia, txt_Year.Text) & "\"
    Set Cn_dbf = New ADODB.Connection
    Cn_dbf.ConnectionString = Cadena
    Cn_dbf.Open
    Cadena = "SELECT *FROM IMPORTAR WHERE YEAR(CGDIA)=" & txt_Year.Text & " AND MONTH(CGDIA)=" & Cbo.ListIndex + 1 '& " ORDER BY CGCIA, CGCOD"
    Set rs = OpenRecordset(Cadena, Cn_dbf)
    Set dg.DataSource = Nothing
    Set dg.DataSource = rs
    lbl(2).Caption = rs.RecordCount
    
    Dim rsTemp As ADODB.Recordset
    Cadena = "SELECT SUM(CGIMP) AS DEBE FROM IMPORTAR WHERE CGMOV = '1' AND YEAR(CGDIA)=" & txt_Year.Text & " AND MONTH(CGDIA)=" & Cbo.ListIndex + 1
    Set rsTemp = OpenRecordset(Cadena, Cn_dbf)
    lbl(4).Caption = IIf(IsNull(rsTemp!Debe), 0, rsTemp!Debe)
    lbl(4).Caption = Round(lbl(4).Caption, 2)
    rsTemp.Close
    Cadena = "SELECT SUM(CGIMP) AS HABER FROM IMPORTAR WHERE CGMOV = '2' AND YEAR(CGDIA)=" & txt_Year.Text & " AND MONTH(CGDIA)=" & Cbo.ListIndex + 1
    Set rsTemp = OpenRecordset(Cadena, Cn_dbf)
    lbl(5).Caption = IIf(IsNull(rsTemp!Haber), 0, rsTemp!Haber)
    lbl(5).Caption = Round(lbl(5).Caption, 2)
    rsTemp.Close
    Set rsTemp = Nothing
End Sub

Private Sub Init_Form()
    txt_Year.Text = Year(Now)
    Me.Top = 0
    Me.Left = 0
    With pbC
        .Min = 0
        .Max = 100
        .Value = 100
    End With
       With pbD
        .Min = 0
        .Max = 100
        .Value = 100
    End With
End Sub

Private Sub Insert()
    Dim MyCon As ADODB.Connection
    Dim Cabecera As String
    Dim Detalle As String
    Dim Data_Source As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim rsSum As ADODB.Recordset
    Dim Ors As ADODB.Recordset
    Dim mVoucher, mLote As String
    Dim i As Integer
    Dim Item As String
    Dim Debe As Double
    Dim Haber As Double
    Dim Glosa As String
On Error GoTo MyErr
    Data_Source = "BDCONT" & txt_Year.Text & ".mdb"
    Set MyCon = New ADODB.Connection
    MyCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & Path_CIA & Cia_StarSoft(wcia, txt_Year.Text) & "\" & Data_Source & ""
    MyCon.Open
    
    Cabecera = "CABMOV" & Format(Cbo.ListIndex + 1, "00")
    Detalle = "DETMOV" & Format(Cbo.ListIndex + 1, "00")
    
    Cadena = "SELECT DISTINCT CGVOUCH FROM IMPORTAR WHERE YEAR(CGDIA)=" & txt_Year.Text & " AND MONTH(CGDIA)=" & Cbo.ListIndex + 1
    Set rsTmp = OpenRecordset(Cadena, Cn_dbf)
    pbC.Max = rsTmp.RecordCount
    pbC.Value = 0
    
    Do While Not rsTmp.EOF
        'INSERTAMOS LA CABECERA
        Cadena = "SELECT TOP 1 *FROM IMPORTAR WHERE YEAR(CGDIA)=" & txt_Year.Text & " AND MONTH(CGDIA)=" & Cbo.ListIndex + 1 & " AND CGVOUCH = '" & rsTmp!CGVOUCH & "'"
        Set rsTemp = OpenRecordset(Cadena, Cn_dbf)

        Do While Not rsTemp.EOF
            Debe = 0
            Haber = 0
            Glosa = Space(1) & "/" & Space(1) & rsTmp!CGVOUCH & Space(1) & rsTemp!CGLOTE
            
            Cadena = "INSERT INTO " & Cabecera & " (SUBDIAR_CODIGO, CMOV_C_COMPR, CMOV_FECHA, CMOV_MONED, CMOV_CONVE, CMOV_DEBE, CMOV_HABER, CMOV_GLOSA) " & _
            "VALUES ('06','" & rsTmp!CGVOUCH & "',#" & Format(rsTemp!CGDIA, "mm/dd/yyyy") & "#,'MN','ESP'," & Debe & "," & Haber & ",'" & Glosa & "')"

            MyCon.Execute Cadena
            
                'INSERTAMOS EL DETALLE
                Cadena = "SELECT *FROM IMPORTAR WHERE YEAR(CGDIA)=" & txt_Year.Text & " AND MONTH(CGDIA)=" & Cbo.ListIndex + 1 & " AND CGVOUCH = '" & rsTmp!CGVOUCH & "'"
                Set Ors = OpenRecordset(Cadena, Cn_dbf)
                i = 1
                pbD.Value = 0
                pbD.Max = Ors.RecordCount
                Do While Not Ors.EOF
                    If i = 572 Then
                        Dim A
                        A = A
                    End If
                    Item = Format(i, "0000")
                    Debe = IIf(Ors!CGMOV = 1, Ors!CGIMP, 0)
                    Haber = IIf(Ors!CGMOV = 2, Ors!CGIMP, 0)
                    Glosa = Space(1) & "/" & Space(1) & Ors!CGVOUCH & Space(1) & Ors!CGLOTE
                    
                    Cadena = "INSERT INTO " & Detalle & " (SUBDIAR_CODIGO, DMOV_C_COMPR, DMOV_SECUE, DMOV_FECHA, DMOV_CUENT, DMOV_FECDC, DMOV_DEBE, DMOV_HABER, DMOV_GLOSA, DMOV_CENCO) " & _
                    "VALUES ('06','" & rsTmp!CGVOUCH & "','" & Item & "',#" & Format(Ors!CGDIA, "mm/dd/yyyy") & "#,'" & Ors!CGCOD & "',#" & Format(Ors!CGDIA, "mm/dd/yyyy") & "#," & Debe & "," & Haber & ",'" & Glosa & "','" & IIf(IsNull(Ors!CENCOS), " ", Ors!CENCOS) & "')"
                    MyCon.Execute Cadena
                    i = i + 1
                    pbD.Value = pbD.Value + 1
                    Ors.MoveNext
                Loop
                
                Cadena = "SELECT SUM(DMOV_DEBE) AS DEBE, SUM(DMOV_HABER) AS HABER FROM " & Detalle & " WHERE YEAR(DMOV_FECHA)=" & txt_Year.Text & " AND MONTH(DMOV_FECHA)=" & Cbo.ListIndex + 1 & " AND DMOV_C_COMPR = '" & rsTemp!CGVOUCH & "' AND SUBDIAR_CODIGO = '06'"
                Set rsSum = OpenRecordset(Cadena, MyCon)
                Debe = IIf(IsNull(rsSum!Debe), 0, rsSum!Debe)
                Haber = IIf(IsNull(rsSum!Haber), 0, rsSum!Haber)
                Cadena = "UPDATE " & Cabecera & " SET CMOV_DEBE = " & Debe & ", CMOV_HABER = " & Haber & " WHERE YEAR(CMOV_FECHA)=" & txt_Year.Text & " AND MONTH(CMOV_FECHA)=" & Cbo.ListIndex + 1 & " AND CMOV_C_COMPR = '" & rsTemp!CGVOUCH & "' AND SUBDIAR_CODIGO = '06'"
                MyCon.Execute Cadena
                
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        pbC.Value = pbC.Value + 1
        rsTmp.MoveNext
    Loop
    If Not MyCon Is Nothing Then
        If rsTemp.State = adStateOpen Then rsTemp.Close
        If rsTmp.State = adStateOpen Then rsTmp.Close
        If MyCon.State = adStateOpen Then MyCon.Close
        Set rsTemp = Nothing
        Set rsTmp = Nothing
        Set MyCon = Nothing
    End If
    MsgBox "Proceso Concluyo Satisfactoriamente", vbInformation + vbOKOnly, "Sistema"
    Unload Me
MyErr:
    If Err.Number <> 0 Then
        MsgBox Err.Number & Space(1) & Err.Description, vbCritical + vbOKOnly, "Sistema"
        Err.Clear
    End If
End Sub


