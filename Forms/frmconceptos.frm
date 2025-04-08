VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmconceptos 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Mantenimiento de Conceptos «"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmconceptos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid dgrconceptos 
      Height          =   4845
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   8546
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   16
      TabAcrossSplits =   -1  'True
      TabAction       =   2
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
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "descripcion"
         Caption         =   "Descripcion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descabrev"
         Caption         =   "Desc. Abreviada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "estado"
         Caption         =   "estado"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   3720.189
         EndProperty
         BeginProperty Column01 
            WrapText        =   -1  'True
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   5010
      Left            =   45
      Top             =   45
      Width           =   6360
   End
End
Attribute VB_Name = "frmconceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsconceptos As New ADODB.Recordset
Dim codcpto_delete As String

Private Sub CreaRepositorio()
  If rsconceptos.State = 1 Then rsconceptos.Close
    rsconceptos.Fields.Append "descripcion", adVarChar, 150, adFldIsNullable
    rsconceptos.Fields.Append "descabrev", adVarChar, 15, adFldIsNullable
    rsconceptos.Fields.Append "estado", adChar, 1, adFldIsNullable
    rsconceptos.Fields.Append "codigo", adVarChar, 3, adFldIsNullable
    
    rsconceptos.Open
    Set dgrconceptos.DataSource = rsconceptos
End Sub

Private Sub dgrconceptos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If OldValue <> dgrconceptos.Columns(ColIndex) Then
    If Len(Trim(dgrconceptos.Columns(2).Value & "")) = 0 Or dgrconceptos.Columns(2).Value = "N" Then dgrconceptos.Columns(2).Value = "N"
    If dgrconceptos.Columns(2).Value = "C" Then dgrconceptos.Columns(2).Value = "M"
End If
End Sub

Private Sub dgrconceptos_BeforeDelete(Cancel As Integer)
codcpto_delete = codcpto_delete & "'" & rsconceptos!codigo & "',"
End Sub

Private Sub dgrconceptos_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub dgrconceptos_OnAddNew()
rsconceptos.AddNew
dgrconceptos.Columns(2).Value = ""
End Sub

Private Sub dgrconceptos_RowResize(Cancel As Integer)
Cancel = True
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
CreaRepositorio
Procesa
End Sub

Private Sub Procesa()
Dim sSQL As String
Dim rs As ADODB.Recordset

codcpto_delete = ""

sSQL = " SELECT cod_concepto,rtrim(desc_concepto) as desc_concepto,desc_abrev_cpto  FROM tconceptos WHERE status!='*' "

Set rs = cn.Execute(sSQL)

If Not rs.EOF Then
    Do While Not rs.EOF
            rsconceptos.AddNew
            
            rsconceptos!Descripcion = rs!desc_concepto
            rsconceptos!descabrev = rs!desc_abrev_cpto
            rsconceptos!estado = "C"
            rsconceptos!codigo = rs!cod_concepto
            
        rs.MoveNext
    Loop
    rs.Close
End If

Set rs = Nothing

End Sub

Public Sub GrabaConceptos()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim NewCodigo As String

rsconceptos.MoveFirst

On Error GoTo GrabaConceptos

cn.BeginTrans

Do While Not rsconceptos.EOF
    If dgrconceptos.Columns(2) = "M" Or dgrconceptos.Columns(2) = "N" Then
        If dgrconceptos.Columns(2) = "N" Then
            sSQL = "SELECT COALESCE(MAX(cod_concepto),0)+1 FROM tconceptos"
            Set rs = cn.Execute(sSQL)
            
            If Not rs.EOF Then
                NewCodigo = Format(rs(0), "000")
                rs.Close
            End If
            Set rs = Nothing
            
            sSQL = "INSERT tconceptos VALUES ('" & NewCodigo & "','" & rsconceptos!Descripcion & "','" & rsconceptos!descabrev & "',' ',GETDATE(),'" & wuser & "',NULL,NULL)"
            
        Else
            sSQL = "UPDATE tconceptos SET desc_concepto='" & rsconceptos!Descripcion & "',desc_abrev_cpto='" & rsconceptos!descabrev & "' WHERE cod_concepto ='" & rsconceptos!codigo & "'"
        End If
        
        cn.Execute sSQL
        
    End If
    rsconceptos.MoveNext
Loop

If Len(Trim(codcpto_delete)) > 1 Then
    codcpto_delete = Mid(codcpto_delete, 1, Len(Trim(codcpto_delete)) - 1)
    
    sSQL = "UPDATE tconceptos SET status='*',fec_modi=getdate(),user_modi='" & wuser & "' WHERE cod_concepto in (" & codcpto_delete & ")"
    cn.Execute (sSQL)
    
End If

cn.CommitTrans

MsgBox "Grabacion Exitosa", vbQuestion, TitMsg
Procesa
Exit Sub

GrabaConceptos:
cn.RollbackTrans
MsgBox "Error " & Err.Description, vbCritical, TitMsg

End Sub
