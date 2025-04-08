VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmRemunAcumulada 
   Caption         =   "Remuneracion Acumulada"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5220
   Begin VB.ComboBox Cmbtipotrabajador 
      Height          =   315
      ItemData        =   "FrmRemunAcumulada.frx":0000
      Left            =   1440
      List            =   "FrmRemunAcumulada.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ver Reporte"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   3975
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
         TabIndex        =   2
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1125
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   255
      Size            =   "450;564"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   345
   End
End
Attribute VB_Name = "FrmRemunAcumulada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlinea As Integer
Dim vAfp As String
Dim mciadir As String
Dim mciatlf As String

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
End Sub

Private Sub Command1_Click()
Call Procesa_remuneracion_Anual
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5295
Me.Height = 2280
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
Txtano.Text = Format(Year(Date), "0000")
End Sub
Public Sub Procesa_remuneracion_Anual()

'Excel
'Dim mcad As String
Dim nFil As Integer
Dim nCol As Integer
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Cells(1, 1).Value = CmbCia
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(1, 1).Font.Size = 12
xlSheet.Cells(1, 1).HorizontalAlignment = xlCenter

xlSheet.Cells(3, 1).Value = "REMUNERACIONES ACUMULADAS"
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).Font.Size = 12
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter

'If Txtsemana.Text = "" Then
'   xlSheet.Cells(4, 1).Value = Space(5) & Cmbmes.Text & Space(5) & Txtano.Text
'Else
'   xlSheet.Cells(4, 1).Value = Space(5) & "SEMANA : " & Txtsemana.Text & Space(5) & Cmbmes.Text & Space(5) & Txtano.Text
'End If
xlSheet.Cells(4, 1).Font.Bold = True
xlSheet.Cells(4, 1).Font.Size = 12
xlSheet.Cells(4, 1).HorizontalAlignment = xlCenter

'xlSheet.Cells(6, 1).Value = "MES"
'xlSheet.Cells(6, 1).Font.Bold = True
'xlSheet.Cells(6, 1).Font.Size = 12
'xlSheet.Cells(6, 1).HorizontalAlignment = xlCenter

'xlSheet.Cells(6, 2).Value = "INGRESOS"
'xlSheet.Cells(6, 2).Font.Bold = True
'xlSheet.Cells(6, 2).Font.Size = 12
'xlSheet.Cells(6, 2).HorizontalAlignment = xlCenter

'xlSheet.Cells(6, 3).Value = "QUINTA"
'xlSheet.Cells(6, 3).Font.Bold = True
'xlSheet.Cells(6, 3).Font.Size = 12
'xlSheet.Cells(6, 3).HorizontalAlignment = xlCenter

nCol = 1
nFil = 8

'-----------------------------------------------------------
Dim rs2 As ADODB.Recordset
Dim mcad As String
Dim mItem As Integer
Dim mtot As Currency
Dim mtot1 As Currency
Dim mtot2 As Currency
Dim mtot3 As Currency
Dim mtot4 As Currency
Dim mtot5 As Currency
Dim mtot6 As Currency
Dim mtot7 As Currency
Dim mtot8 As Currency
Dim mtot9 As Currency
Dim mtot10 As Currency
Dim mtot11 As Currency
Dim mtot12 As Currency
Dim mtot13 As Currency
Dim MUIT As Currency

mciadir = ""
 
Sql$ = "select direcc,nro,dpto from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then mciadir = Trim(rs!direcc) & " " & Trim(rs!NRO)
If rs.State = 1 Then rs.Close
Sql$ = "select telef from telef_cia where cod_cia='" & wcia & "'"
If (fAbrRst(rs, Sql$)) Then mciatlf = Trim(rs!telef)
If rs.State = 1 Then rs.Close

          
Sql$ = "EXEC PLANSS_REMUNERACION_ACUMULADA '" & wcia & "','" & Txtano.Text & "','" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "'"

     
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Existen Datos Registrados", vbInformation, "Liquidacion Anual de Afp": Exit Sub
rs.MoveFirst
mItem = 1

RUTA$ = App.Path & "\REPORTS\" & "REMUN.txt"

Open RUTA$ For Output As #1

mtot = 0: mtot1 = 0: mtot2 = 0: mtot3 = 0: mtot4 = 0: mtot5 = 0: mtot6 = 0: mtot7 = 0
mtot8 = 0: mtot9 = 0: mtot10 = 0: mtot11 = 0: mtot12 = 0: mtot13 = 0:

Print #1, LetraChica

Call Cabeza_remuneracion_Anual
 '
xlSheet.Cells(nFil, 1).Value = "No  Codigo   Nombre"
xlSheet.Cells(nFil, 2).Value = "Remun "
xlSheet.Cells(nFil, 3).Value = "Util  "
xlSheet.Cells(nFil, 4).Value = "Inc.Afp"
xlSheet.Cells(nFil, 5).Value = "Gratif."
xlSheet.Cells(nFil, 6).Value = "Ing.Total"
xlSheet.Cells(nFil, 7).Value = "AFP 3%  "
xlSheet.Cells(nFil, 8).Value = "Rem.Qta "
xlSheet.Cells(nFil, 9).Value = "7 UIT "
xlSheet.Cells(nFil, 10).Value = "Remun. Afecta"
xlSheet.Cells(nFil, 11).Value = "Impuesto Calcul."
xlSheet.Cells(nFil, 12).Value = "Impuesto Retenido"
xlSheet.Cells(nFil, 13).Value = "Diferencia "
xlSheet.Cells(nFil, 14).Value = "F.Cese"


Do While Not rs.EOF
   Sql$ = "select ap_pat,ap_mat,nom_1,nom_2,fcese from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
   If Not (fAbrRst(rs2, Sql$)) Then
      MsgBox "No se Encuentra Registrado el Codigo de Trabajador", vbCritical, "Codigo => " & rs!PlaCod
      If rs.State = 1 Then rs.Close
      If rs2.State = 1 Then rs2.Close
      Exit Sub
   End If
   mcad = ""

   'MODIFICADO EL 08/01/2008
   NUMAFP = IIf(IsNull(rs!PlaCod), rs!PlaCod, rs!PlaCod)
   NOM$ = rs!nombres
   mcad = fCadNum(mItem, "###") & " " & lentexto(8, Left(NUMAFP, 8)) & " " & lentexto(28, Left(NOM$, 28))
   nFil = nFil + 1

   For I = 2 To 13
   
       If rs(I) = 0 Then
          If I = 2 Or I = 6 Or I = 8 Or I = 9 Or I = 13 Then
            mcad = mcad & Space(13)
          Else
            mcad = mcad & Space(11)
          End If
       Else
          If I = 2 Or I = 6 Or I = 8 Or I = 9 Or I = 13 Then
            mcad = mcad & Space(1) & fCadNum(rs(I), "#,###,###.00")
          Else
            mcad = mcad & Space(1) & fCadNum(rs(I), "###,###.00")
          End If
       End If
       
       nCol = 1
       xlSheet.Cells(nFil, nCol).Value = fCadNum(mItem, "###") & " " & lentexto(8, Left(NUMAFP, 8)) & " " & lentexto(28, Left(NOM$, 28))
       nCol = I
       xlSheet.Cells(nFil, nCol).Value = fCadNum(rs(I), "##,###,###.00")
       
       Select Case I
              Case Is = 2
                    mtot1 = mtot1 + rs(I)
                    
              Case Is = 3
                    mtot2 = mtot2 + rs(I)
                   
              Case Is = 4
                    mtot3 = mtot3 + rs(I)
                   
              Case Is = 5
                    mtot4 = mtot4 + rs(I)
                   
              Case Is = 6
                    mtot5 = mtot5 + rs(I)
                   
              Case Is = 7
                    mtot6 = mtot6 + rs(I)
                   
              Case Is = 8
                    mtot7 = mtot7 + rs(I)
                   
              Case Is = 9
                    mtot8 = mtot8 + rs(I)
                    
              Case Is = 10
                    mtot9 = mtot9 + rs(I)
                    
              Case Is = 11
                    mtot10 = mtot10 + rs(I)
                    
              Case Is = 12
                    mtot11 = mtot11 + rs(I)
                    
              Case Is = 13
                    mtot12 = mtot12 + rs(I)
                   
         End Select
       
         If I >= 13 Then Exit For
         
   Next I

  
   mcad = mcad & Space(2) & Format(rs2!fcese, "dd/mm/yyyy")
   
   xlSheet.Cells(nFil, 14).Value = Format(rs2!fcese, "dd/mm/yyyy")
   
   Print #1, mcad
   
   If rs2.State = 1 Then rs2.Close
   
   rs.MoveNext
   
   mItem = mItem + 1
   
   mlinea = mlinea + 1
   
   If mlinea > 55 Then Print #1, SaltaPag: Cabeza_remuneracion_Anual
   
  mtot = 0
  
Loop

If rs.State = 1 Then rs.Close

Print #1, Space(42) & String(153, "-")

nFil = nFil + 1
nCol = 4

'EXCEL
xlSheet.Cells(nFil, nCol).Value = fCadNum(mtot1, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 1).Value = fCadNum(mtot2, "#,###,###.00")
xlSheet.Cells(nFil, nCol + 2).Value = fCadNum(mtot3, "#,###,###.00")
xlSheet.Cells(nFil, nCol + 3).Value = fCadNum(mtot4, "#,###,###.00")
xlSheet.Cells(nFil, nCol + 4).Value = fCadNum(mtot5, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 5).Value = fCadNum(mtot6, "#,###,###.00")
xlSheet.Cells(nFil, nCol + 6).Value = fCadNum(mtot7, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 7).Value = fCadNum(mtot8, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 8).Value = fCadNum(mtot9, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 9).Value = fCadNum(mtot10, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 10).Value = fCadNum(mtot11, "#,###,###.00")
xlSheet.Cells(nFil, nCol + 11).Value = fCadNum(mtot12, "##,###,###.00")
xlSheet.Cells(nFil, nCol + 12).Value = fCadNum(mtot13, "###,###.00")

mcad = Space(42) & fCadNum(mtot1, "##,###,###.00") & Space(1) & fCadNum(mtot2, "#,###,###.00") & Space(1) & fCadNum(mtot3, "#,###,###.00") & Space(1) & fCadNum(mtot4, "#,###,###.00") & Space(1) & fCadNum(mtot5, "##,###,###.00") & Space(1)
mcad = mcad & fCadNum(mtot6, "#,###,###.00") & Space(1) & fCadNum(mtot7, "##,###,###.00") & Space(1) & fCadNum(mtot8, "##,###,###.00") & Space(1) & fCadNum(mtot9, "##,###,###.00") & Space(1) & fCadNum(mtot10, "##,###,###.00") & Space(1)
mcad = mcad & fCadNum(mtot11, "#,###,###.00") & Space(1) & fCadNum(mtot12, "##,###,###.00") '& Space(1) & fCadNum(mtot13, "###,###.00")

Print #1, mcad
Close #1

xlSheet.Range("A:AZ").EntireColumn.AutoFit

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "REMUNERACIONES ACUMULADAS"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = vbDefault


Call Imprime_Txt("REMUN.txt", RUTA$)
End Sub
Private Sub Cabeza_remuneracion_Anual()

Print #1, "Empresa     " & CmbCia.Text & Space(5) & "Ruc   :  " & wruc
Print #1, "Direccion   " & mciadir & "             Telefono  : " & mciatlf
Print #1, String(195, "-")
Print #1, "REPORTE DE REMUNERACION ACUMULADA " & " " & Txtano.Text
Print #1,
'Print #1, "AFP         " & Cmbafp.Text
Print #1,
Print #1, " No  Codigo   Nombre                         Remun          Util     Inc.Afp    Gratif.    Ing.Total     AFP 3%      Rem.Qta         7UIT     Remun.   Impuesto   Impuesto  Diferencia     F.Cese"
Print #1, "                                                                                                                                              Afecta    Calcul.   Retenido"
Print #1, String(195, "-")
mlinea = 11
End Sub
Private Sub SpinButton1_SpinDown()
If Val(Txtano.Text) > 0 Then Txtano = Txtano - 1
End Sub
Private Sub SpinButton1_SpinUp()
If Val(Txtano) = 0 Then Txtano = "1" Else Txtano = Txtano + 1
End Sub
Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Procesa_remuneracion_Anual
End Sub
