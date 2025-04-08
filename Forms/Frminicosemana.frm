VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frminicosemana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio Anual de semana"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "Frminicosemana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5655
      Begin VB.TextBox Txtinicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Text            =   "1"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Txtano 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin MSMask.MaskEdBox Txtfecha 
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSForms.SpinButton SpinButton2 
         Height          =   300
         Left            =   4800
         TabIndex        =   10
         Top             =   600
         Width           =   255
         Size            =   "450;529"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dia de Inicio de Mes"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio de Semana"
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   80
         Width           =   4335
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "Frminicosemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 2190
Me.Width = 6060
Txtano.Text = Format(Year(Date) + 1, "0000")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Sql$ = "select iniciomes from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   If IsNull(rs!iniciomes) Then
      Txtinicio.Text = ""
   Else
      Txtinicio.Text = rs!iniciomes
   End If
End If
rs.Close
End Sub
Public Sub Grabar_InicioSem()
Dim FecFin As String
Dim mf As String
Dim NroTrans As Integer
'On Error GoTo ErrorTrans
NroTrans = 0

'***********codigo agregado giovanni 1392007******************************
Dim MesCalculado As String
Dim TipoCalculo As String
Dim VueltasMes As Integer
MesCalculado = "": TipoCalculo = "N": VueltasMes = 0
'*************************************************************************

If Not IsDate(Txtfecha) Then
   MsgBox "Ingrese Correctamente la Fecha", vbCritical, "Inicio de Planilla"
   Txtfecha.SetFocus
   Exit Sub
End If

If Not IsNumeric(Txtano) Then
   MsgBox "Ingrese Correctamente Año", vbCritical, "Inicio de Planilla"
   Txtano.SetFocus
   Exit Sub
End If

If Not IsNumeric(Txtinicio) Then
   MsgBox "Ingrese Correctamente Dia de Inicio", vbCritical, "Inicio de Planilla"
   Txtinicio.SetFocus
   Exit Sub
End If

Mgrab = MsgBox("Seguro de Grabar Inicio de Semana", vbYesNo + vbQuestion, "Inicio de Semana")
If Mgrab <> 6 Then Exit Sub

cn.BeginTrans
NroTrans = 1

If wGrupoPla = "01" Then 'Para Grupo Gallos se Paga Asig. Fam. en la ultima semana del mes
   Dim Rq As ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*' and prefijo<>'' and not prefijo is null"
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
      Sql$ = "update plasemanas set status='*' where cia='" & Rq!cod_cia & "' and ano='" & Trim(Txtano.Text) & "' and status<>'*'"
      cn.Execute Sql
      mf = Txtfecha.Text
      For I = 1 To 53
          FecFin = Calcula_Fecha_final(mf)
          If Month(Calcula_Fecha_final(Calcula_Fecha_Inicial(FecFin))) <> Month(FecFin) Then
              TipoCalculo = "S"
          Else
             TipoCalculo = "N"
          End If
          Sql$ = "SET DATEFORMAT MDY insert plasemanas values('" & Rq!cod_cia & "','" & Format(I, "00") & "','" & Format(mf, FormatFecha) & "','" & Format(FecFin, FormatFecha) & "','" & Txtano.Text & "','','','" & wuser & "'," & FechaSys & ",'" & TipoCalculo & "')"
          cn.Execute Sql
          mf = Calcula_Fecha_Inicial(FecFin)
      Next I
      Sql$ = "update cia set iniciomes='" & Trim(Txtinicio.Text) & "' where cod_cia='" & Trim(Rq!cod_cia) & "' and status<>'*'"
      cn.Execute Sql
      Rq.MoveNext
   Loop
   Rq.Close: Set Rq = Nothing
Else
   Sql$ = "update plasemanas set status='*' where cia='" & wcia & "' and ano='" & Trim(Txtano.Text) & "' and status<>'*'"
   cn.Execute Sql
   mf = Txtfecha.Text
   For I = 1 To 53
       FecFin = Calcula_Fecha_final(mf)
       '**********codigo agregado giovanni 13092007*****************************
       If MesCalculado <> Month(FecFin) Then
           VueltasMes = 1: MesCalculado = Month(FecFin)
       Else
           VueltasMes = VueltasMes + 1
           If VueltasMes = 2 Then
               TipoCalculo = "S"
           Else
               TipoCalculo = "N"
           End If
       End If
    '************************************************************************
       Sql$ = "SET DATEFORMAT MDY insert plasemanas values('" & wcia & "','" & Format(I, "00") & "','" & Format(mf, FormatFecha) & "','" & Format(FecFin, FormatFecha) & "','" & Txtano.Text & "','','','" & wuser & "'," & FechaSys & ",'" & TipoCalculo & "')"
       cn.Execute Sql
       mf = Calcula_Fecha_Inicial(FecFin)
   Next I
   Sql$ = "update cia set iniciomes='" & Trim(Txtinicio.Text) & "' where cod_cia='" & Trim(wcia) & "' and status<>'*'"
   cn.Execute Sql
End If


Sql$ = "update plasemanas set status='*' where cia='" & wcia & "' and ano='" & Trim(Txtano.Text) & "' and status<>'*' and year(fechaf)>convert(int,ano)"
cn.Execute Sql

cn.CommitTrans
MsgBox "Proceso concluido Satisfactoriamente", vbInformation
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
End Sub
Private Function Calcula_Fecha_final(fecha As String) As String
Dim mdia As Integer
Dim mmes As Integer
Dim Mes As Integer
Dim mano As Integer
Dim MaxDia As Integer
mdia = Val(Mid(fecha, 1, 2))
mmes = Val(Mid(fecha, 4, 2))
mano = Val(Mid(fecha, 7, 4))
Mes = Val(Mid(fecha, 4, 2))
Select Case Mes
       Case Is = 1, 3, 5, 7, 8, 10
            If mdia <= 25 Then
               mdia = mdia + 6
               mmes = mmes
               mano = mano
            Else
               mmes = mmes + 1
                mdia = 6 - (31 - mdia)
               mano = mano
            End If
       Case Is = 2
            If (mano Mod 4) = 0 Then MaxDia = 29 Else MaxDia = 28
            If mdia <= (MaxDia - 6) Then
               mdia = mdia + 6
               mmes = mmes
               mano = mano
            Else
               mmes = mmes + 1
               mdia = 6 - (MaxDia - mdia)
               mano = mano
            End If
       Case Is = 4, 6, 9, 11
            If mdia <= 24 Then
               mdia = mdia + 6
               mmes = mmes
               mano = mano
            Else
               mmes = mmes + 1
               mdia = 6 - (30 - mdia)
               mano = mano
            End If
       Case Is = 12
            If mdia <= 25 Then
               mdia = mdia + 6
               mmes = mmes
               mano = mano
            Else
               mmes = 1
               mdia = 6 - (31 - mdia)
               mano = mano + 1
            End If
End Select
Calcula_Fecha_final = Format(mdia, "00") & "/" & Format(mmes, "00") & "/" & Format(mano, "0000")
End Function
Private Function Calcula_Fecha_Inicial(fecha As String) As String
Dim mdia As Integer
Dim mmes As Integer
Dim Mes As Integer
Dim mano As Integer
Dim MaxDia As Integer
mdia = Val(Mid(fecha, 1, 2))
mmes = Val(Mid(fecha, 4, 2))
mano = Val(Mid(fecha, 7, 4))
Mes = Val(Mid(fecha, 4, 2))
Select Case Mes
       Case Is = 1, 3, 5, 7, 8, 10
            If mdia < 31 Then
               mdia = mdia + 1
               mmes = mmes
               mano = mano
            Else
               mdia = 1
               mmes = mmes + 1
               mano = mano
            End If
       Case Is = 2
            If (mano Mod 4) = 0 Then MaxDia = 29 Else MaxDia = 28
            If mdia < MaxDia Then
               mdia = mdia + 1
               mmes = mmes
               mano = mano
            Else
               mdia = 1
               mmes = mmes + 1
               mano = mano
            End If
       Case Is = 4, 6, 9, 11
            If mdia < 30 Then
               mdia = mdia + 1
               mmes = mmes
               mano = mano
            Else
               mdia = 1
               mmes = mmes + 1
               mano = mano
            End If
       Case Is = 12
            If mdia < 31 Then
               mdia = mdia + 1
               mmes = mmes
               mano = mano
            Else
               mdia = 1
               mmes = 1
               mano = mano + 1
            End If
End Select
Calcula_Fecha_Inicial = Format(mdia, "00") & "/" & Format(mmes, "00") & "/" & Format(mano, "0000")
End Function

Private Sub Label5_Click()

End Sub

Private Sub SpinButton2_SpinDown()
If Val(Txtinicio.Text) > 1 Then
   Txtinicio.Text = Val(Txtinicio.Text) - 1
Else
   Txtinicio.Text = "1"
End If
End Sub

Private Sub SpinButton2_SpinUp()
If Val(Txtinicio.Text) < 30 Then
   Txtinicio.Text = Val(Txtinicio.Text) + 1
Else
   Txtinicio.Text = "30"
End If
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtfecha.SetFocus
End Sub
Private Sub Txtano_LostFocus()
Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Trim(Txtano.Text) & "' and semana='01' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   Txtfecha.Text = Format(rs!fechai, "dd/mm/yyyy")
Else
    Txtfecha.Text = "__/__/____"
End If
End Sub

Private Sub Txtinicio_KeyPress(KeyAscii As Integer)
Txtinicio.Text = Txtinicio.Text + fc_ValNumeros(KeyAscii)
End Sub
