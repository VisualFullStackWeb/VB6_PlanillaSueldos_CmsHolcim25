VERSION 5.00
Begin VB.Form frmactpres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Cantidad"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmactpres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkgrati 
         Alignment       =   1  'Right Justify
         Caption         =   "Dscto Grati"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3555
         MaskColor       =   &H8000000F&
         TabIndex        =   7
         Top             =   585
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   5085
         MaskColor       =   &H8000000F&
         TabIndex        =   6
         Top             =   555
         Width           =   1380
      End
      Begin VB.CommandButton cmd_act 
         Caption         =   "&Actualizar"
         Height          =   345
         Left            =   1920
         MaskColor       =   &H8000000F&
         TabIndex        =   5
         Top             =   555
         Width           =   1380
      End
      Begin VB.TextBox txtcuota 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   560
         Width           =   870
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuota :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblnombre 
         BackStyle       =   0  'Transparent
         Caption         =   "__________"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1815
         TabIndex        =   2
         Top             =   0
         Width           =   4635
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres y Apellidos :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmactpres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fecha As String
Public PlaCod As String
Public codaux As String
Private Sub cmd_act_Click()
On Error GoTo CORRIGE
Dim con As String
Dim NroTrans As Integer
NroTrans = 0
cn.BeginTrans
NroTrans = 1
con = "UPDATE PLACTACTE SET PARTES='" & txtcuota & "',sn_grati=" & IIf(chkgrati.Value = vbChecked, 1, 0) & " WHERE " & _
"PLACOD='" & PlaCod & "' and cia='" & wcia & "' AND id_doc=" & txtcuota.Tag
cn.Execute con
cn.CommitTrans
MsgBox "Datos Actualizados", vbQuestion, Me.Caption
Unload Me
Exit Sub
CORRIGE:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox "Error :" & ERR.Description, vbCritical, Me.Caption
End Sub

Private Sub Command1_Click()
Dim sSQL As String
Dim rsdatos As ADODB.Recordset
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
sSQL = "select importe,pago_acuenta from PLACTACTE WHERE PLACOD='" & PlaCod & "' and cia='" & wcia & "' and status != '*' AND id_doc=" & txtcuota.Tag
Set rsdatos = cn.Execute(sSQL)

If Not rsdatos.EOF Then
    If rsdatos(1) > 0 Then
        MsgBox "No se Puede Anular Prestamo, por tener pagos parciales", vbExclamation, "Prestamos"
        rsdatos.Close
        Set rsdatos = Nothing
        Exit Sub
    End If
    
    rsdatos.Close
End If

Set rsdatos = Nothing

If MsgBox("Seguro de Elimar Prestamos", vbQuestion + vbYesNo, "Prestamos") = vbYes Then
    cn.BeginTrans
    NroTrans = 1
    sSQL = "update plactacte set status='*' where PLACOD='" & PlaCod & "' and cia='" & wcia & "' AND id_doc=" & txtcuota.Tag
    cn.Execute sSQL
    cn.CommitTrans
    MsgBox "Anulación del Prestamo N° : " & txtcuota.Tag & " satisfactoria.", vbInformation, "Prestamos"
    
    Unload Me
End If
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then
           Unload Me
        End If
End Sub

Private Sub Form_Load()
KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Frmgrdctacte.ProcesaNew
End Sub

