VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Cons1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10140
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
      Height          =   495
      Left            =   8745
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   375
      TabIndex        =   1
      Top             =   225
      Width           =   9615
      Begin VB.CommandButton cmd_Consultar 
         Caption         =   "&Consultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7485
         TabIndex        =   8
         Top             =   390
         Width           =   1800
      End
      Begin VB.ComboBox cmb_tt 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   2190
      End
      Begin MSComCtl2.DTPicker dtp_desde 
         Height          =   300
         Left            =   1650
         TabIndex        =   5
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   40490
      End
      Begin MSComCtl2.DTPicker dtp_Hasta 
         Height          =   300
         Left            =   5040
         TabIndex        =   6
         Top             =   195
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   40490
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta"
         Height          =   195
         Left            =   3660
         TabIndex        =   4
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde"
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   240
         Width           =   960
      End
   End
   Begin MSDataGridLib.DataGrid dg_data 
      Height          =   3810
      Left            =   405
      TabIndex        =   0
      Top             =   2205
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   6720
      _Version        =   393216
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Resultados de la Consultas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   435
      TabIndex        =   10
      Top             =   1800
      Width           =   4275
   End
End
Attribute VB_Name = "frm_Cons1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -- Variables para Excel
Dim Obj_Excel   As Object
Dim Obj_Libro   As Object
Dim Obj_Hoja    As Object



Private Sub cmd_Consultar_Click()
On Error GoTo LaCagada

Dim query As String
Dim Rs_tmp As ADODB.Recordset
Dim coman As New ADODB.Command
Dim TT As String


If cmb_tt.Text = "" Then
    MsgBox "Seleccione un tipo de trabajador"
    cmb_tt.SetFocus
    Exit Sub
End If




query = "pl_Rep0001 '" & dtp_desde.Value & "','" & dtp_Hasta.Value & "','" & Trim(Right(cmb_tt.Text, 3)) & "','" & wcia & "'"

TT = Trim(Right(cmb_tt.Text, 3))
coman.CommandText = "pl_Rep0001"
coman.CommandType = adCmdStoredProc
coman.ActiveConnection = cn
coman.Parameters("@f1") = dtp_desde.Value
coman.Parameters("@f2") = dtp_Hasta.Value
coman.Parameters("@tt") = TT
coman.Parameters("@cia") = wcia

Set Rs_tmp = coman.Execute
Set dg_data.DataSource = Rs_tmp
dg_data.Columns(0).Width = 700
dg_data.Columns(1).Width = 4000
dg_data.Columns(2).Width = 1200
dg_data.Columns(3).Width = 900
dg_data.Columns(4).Width = 900

dg_data.Columns(2).Alignment = dbgCenter
dg_data.Columns(3).Alignment = dbgRight
dg_data.Columns(4).Alignment = dbgRight

'


Exit Sub
LaCagada:
MsgBox Err.Description

End Sub

Private Sub Command1_Click()
Call exportar_Datagrid(dg_data.ApproxCount)
End Sub

Private Sub Form_Load()
Call CargarDatos
End Sub

Private Sub CargarDatos()
'
'
Call Llena_Tipo_Trabajadores

End Sub

Sub Llena_Tipo_Trabajadores()
    Call Recupera_Tipos_Trabajadores
    Set rs_MGenera = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    Do While Not rs_MGenera.EOF
        cmb_tt.AddItem rs_MGenera!descrip & Space(50) & rs_MGenera!COD_MAESTRO2
        rs_MGenera.MoveNext
    Loop
    Set rs_MGenera = Nothing
End Sub


Private Sub exportar_Datagrid(n_Filas As Long)
  
   
  
    On Error GoTo Error_Handler
      
    Dim i   As Integer
    Dim j   As Integer
      
    ' -- Colocar el cursor de espera mientras se exportan los datos
    
    Me.MousePointer = vbHourglass
      
    If n_Filas = 0 Then
        MsgBox "No hay datos para exportar a excel. Se ha indicado 0 en el parámetro Filas ": Exit Sub
    Else
          
        ' -- Crear nueva instancia de Excel
        Set Obj_Excel = CreateObject("Excel.Application")
        ' -- Agregar nuevo libro
        'Set Obj_Libro = Obj_Excel.Workbooks.Open(path)
        Set Obj_Libro = Obj_Excel.Workbooks.Add
      
        ' -- Referencia a la Hoja activa ( la que añade por defecto Excel )
        Set Obj_Hoja = Obj_Excel.ActiveSheet
     
        iCol = 0
        ' --  Recorrer el Datagrid ( Las columnas )
        
        For i = 0 To dg_data.Columns.count - 1
            If dg_data.Columns(i).Visible Then
                ' -- Incrementar índice de columna
                iCol = iCol + 1
                ' -- Obtener el caption de la columna
                Obj_Hoja.Cells(1, iCol) = dg_data.Columns(i).Caption
                ' -- Recorrer las filas
                For j = 0 To n_Filas - 1
                    ' -- Asignar el valor a la celda del Excel
                    Obj_Hoja.Cells(j + 2, iCol) = _
                    dg_data.Columns(i).CellValue(dg_data.GetBookmark(j))
                Next
            End If
        Next
          
        ' -- Hacer excel visible
        Obj_Excel.Visible = True
          
        ' -- Opcional : colocar en negrita y de color rojo los enbezados en la hoja
        With Obj_Hoja
            .Rows(1).Font.Bold = True
            .Rows(1).Font.Color = vbRed
            ' -- Autoajustar las cabeceras
            .Columns("A:Z").AutoFit
        End With
    End If
  
    ' -- Eliminar las variables de objeto excel
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
      
    ' -- Restaurar cursor
    Me.MousePointer = vbDefault
      
Exit Sub
  
' -- Error
Error_Handler:
  
    MsgBox Err.Description, vbCritical
    On Error Resume Next
  
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    Me.MousePointer = vbDefault
  
End Sub
