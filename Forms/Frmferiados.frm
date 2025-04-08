VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmferiados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Registro de dias Feriados «"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "Frmferiados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   7935
      Begin VB.TextBox Txtano 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "Frmferiados.frx":030A
         Left            =   1200
         List            =   "Frmferiados.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Width           =   645
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Top             =   120
         Width           =   255
         Size            =   "450;503"
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   1125
      Width           =   7695
      Begin VB.CheckBox Check 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   30
         Left            =   2400
         TabIndex        =   39
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   29
         Left            =   1320
         TabIndex        =   38
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   28
         Left            =   240
         TabIndex        =   37
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   27
         Left            =   6720
         TabIndex        =   36
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   26
         Left            =   5640
         TabIndex        =   35
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   25
         Left            =   4560
         TabIndex        =   34
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   24
         Left            =   3480
         TabIndex        =   33
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   23
         Left            =   2400
         TabIndex        =   32
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   22
         Left            =   1320
         TabIndex        =   31
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   21
         Left            =   240
         TabIndex        =   30
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   20
         Left            =   6720
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   19
         Left            =   5640
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   18
         Left            =   4560
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   17
         Left            =   3480
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   16
         Left            =   2400
         TabIndex        =   25
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   15
         Left            =   1320
         TabIndex        =   24
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   14
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   13
         Left            =   6720
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   12
         Left            =   5640
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   11
         Left            =   4560
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   10
         Left            =   3480
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   9
         Left            =   2400
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   8
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   6
         Left            =   6720
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   5
         Left            =   5640
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   4
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   3
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   2
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   135
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   165
         TabIndex        =   1
         Top             =   135
         Width           =   825
      End
   End
End
Attribute VB_Name = "Frmferiados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mdia As Integer
Dim mfechaI As String
Dim mFecha As String
Dim Sql As String


Private Sub Check_Click(index As Integer)
Call Check_Clic(Check(index))
End Sub

Private Sub Cmbmes_Click()
Validar_UltimoDia
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 7800
Me.Height = 4770
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Cmbmes.ListIndex = Month(Date) - 1
Txtano.Text = Format(Year(Date), "0000")
End Sub
Private Sub Check_Clic(ByRef Chk As Control)
If Chk.Value = 0 Then Chk.ForeColor = &H800000 Else Chk.ForeColor = &HFF&
End Sub
Public Sub Grabar_Feriados()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
Dim Mgrab As Integer
Dim I As Integer
NroTrans = 0
Mgrab = MsgBox("Seguro de Grabar Seteo de Feriados", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub

cn.BeginTrans
NroTrans = 1

Sql = "update plaferiados set status='*' where cia='" & wcia & "' and tipo='FE' " _
    & "and year(fecha)=" & Txtano & " and month(fecha)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
cn.Execute Sql

For I = 0 To mdia - 1
    If Check(I).Value = 1 Then
       mFecha = Format(I + 1, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text
       Sql = "SET DATEFORMAT MDY INSERT INTO plaferiados values('" & wcia & "','FE','" & Format(mFecha, FormatFecha) & "','" & wuser & "',''," & FechaSys & ")"
     
       cn.Execute Sql
    End If
Next

cn.CommitTrans
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption

End Sub
Private Sub Validar_UltimoDia()
Dim I As Integer
If Trim(Txtano.Text) = "" Then Exit Sub
mdia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
For I = 0 To 30
    Check(I).Value = 0
Next
Select Case mdia
       Case Is = 28
             Check(28).Value = 0: Check(28).Visible = False
             Check(29).Value = 0: Check(29).Visible = False
             Check(30).Value = 0: Check(30).Visible = False
       Case Is = 29
             Check(28).Visible = True
             Check(29).Value = 0: Check(29).Visible = False
             Check(30).Value = 0: Check(30).Visible = False
       Case Is = 30
             Check(28).Visible = True
             Check(29).Visible = True
             Check(30).Value = 0: Check(30).Visible = False
       Case Is = 31
             Check(28).Visible = True
             Check(29).Visible = True
             Check(30).Visible = True
End Select

Sql = "select * from plaferiados where cia='" & wcia & "' and tipo='FE' " _
    & "and year(fecha)=" & Txtano & " and month(fecha)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
    
If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Do While Not rs.EOF
   Check(Day(rs!fecha) - 1).Value = 1
   rs.MoveNext
Loop
End Sub

Private Sub Txtano_Change()
Validar_UltimoDia
End Sub
