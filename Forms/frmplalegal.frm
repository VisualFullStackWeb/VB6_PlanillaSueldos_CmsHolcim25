VERSION 5.00
Begin VB.Form frmplalegal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdprint 
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4455
      TabIndex        =   13
      Top             =   45
      Width           =   1050
   End
   Begin VB.PictureBox PctCab 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDDDE&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   135
      ScaleHeight     =   1545
      ScaleWidth      =   4260
      TabIndex        =   2
      Top             =   360
      Width           =   4290
      Begin VB.TextBox TxtSemana 
         Height          =   285
         Left            =   1050
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ComboBox cbotipobol 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmplalegal.frx":0000
         Left            =   1125
         List            =   "frmplalegal.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   450
         Width           =   2985
      End
      Begin VB.ComboBox cbotipotrab 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmplalegal.frx":0034
         Left            =   1125
         List            =   "frmplalegal.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   90
         Width           =   2985
      End
      Begin VB.TextBox txtaño 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         TabIndex        =   4
         Top             =   825
         Width           =   960
      End
      Begin VB.ComboBox cbomes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmplalegal.frx":0056
         Left            =   630
         List            =   "frmplalegal.frx":0081
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   825
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Boleta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2565
         TabIndex        =   7
         Top             =   870
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trab :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lblmes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   870
         Width           =   435
      End
   End
   Begin VB.OptionButton optdetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Impresion Detalle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2475
      TabIndex        =   1
      Top             =   45
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.OptionButton optcab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Impresion Cabezera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   2130
   End
   Begin VB.PictureBox pctDet 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDDDE&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   135
      ScaleHeight     =   1470
      ScaleWidth      =   4260
      TabIndex        =   8
      Top             =   360
      Width           =   4290
      Begin VB.ComboBox cbobol 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmplalegal.frx":00EA
         Left            =   1260
         List            =   "frmplalegal.frx":00F7
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   450
         Width           =   2805
      End
      Begin VB.ComboBox cbotrab 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmplalegal.frx":011E
         Left            =   1260
         List            =   "frmplalegal.frx":0128
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   90
         Width           =   2805
      End
      Begin VB.TextBox txtinicio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   12
         Top             =   1080
         Width           =   1680
      End
      Begin VB.TextBox txtfin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2295
         TabIndex        =   11
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Boleta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trab :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   135
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Fin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   10
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblinicio 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   585
         TabIndex        =   9
         Top             =   810
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmplalegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim mtipo As String
'Dim rs2 As ADODB.Recordset
Dim s_Tipo_Trabajador As String
Dim s_Tipo_Boleta As String
Dim i_Contador_Registros_2 As Integer
Private Sub Cabeza_Resumen()
Dim rsremu As ADODB.Recordset
Dim mcadir As String
Dim mcadinr As String
Dim mcadd As String
Dim mcada As String
Dim mcadh As String
Dim cad As String
Dim i As Long
Dim RUTA$

Dim MCADENA As String
RUTA$ = App.path & "\REPORTS\" & "REPORTELEGAL.txt"

Open RUTA$ For Output As #1
For i = Val(txtinicio.Text) To Val(txtfin.Text)
    If i <> Val(txtinicio.Text) Then
        Print #1, Chr(12) + Chr(13)
    End If
    Print #1, Chr(18)
    Print #1, "RODA S.A."
    Print #1, Chr(14)
    Print #1, Space(40) & Chr(14) & "PLANILLA DE PAGO " & cbotrab.Text & Space(10) & Format(i, "0000") & Chr(20)
    Print #1, Chr(15)
    Print #1, String(233, "-")
    Print #1, Chr(169) & String(49, "-") & "     DATOS DEL TRABAJADOR     " & String(41, "-") & Chr(170) & " " & Chr(169) & String(38, "-") & "     DATOS DEL TRABAJADOR     " & String(35, "-") & Chr(172)
    Print #1, "Codigo   Apellidos y Nombres del Trabajador        Ocupacion         F. Ingreso   F.cese   Carnet I.P.S.S.     Jornal    * *- H.Trab.   H.domi.  H.Feri   H.E.L-S 2P H.E.D-F  H.E.N.O.  H.B.BProd.  H.Vacac.  H.Des.Med  H.E.L-S 3era"
    Print #1, Space(70) & " INGRESOS REMUNERATIVOS "
    Print #1, "BASICO     DOMINICAL  I.H.E.L-S  ASIG. FAM. REM.FERIA. B.COS.VIDA OTROS PAGOS VACACIO.  GRATIFI.  REINTEGRO  B.H.E. D-F  B.H.E N-O  A.F.P.  B.T.SERV.  HOR.VERA.  PRIMA TUR  PRI.PROD  BON.PROD. DESC.MED.  BON.RENDI.  I.E.L-S 3era"
    Print #1, Space(15) & " INGRESOS NO REMUNERATIVOS "
    Print #1, "MOVILIDAD  VIATICOS   ASIG.ESC.  PAR.UTIL. COND.TRAB D.TRAB."
    Print #1, String(55, "-") & "  DEDUCCIONES  " & String(67, "-") & " " & String(30, "-") & "  APORTACIONES  " & String(30, "-")
    Print #1, "ESSALUD    S.N.P.    I-E-S      SENATI      SCTR-S     ADELANTOS  RET.JUDIC.  CTA.CTE.   OTR.DESC.  QTA.CAT.    A.F.P.     ESSALUD-VIDA ***ESSALUD    S.N.P.    I-E-S   SENATI   SCTR-S    *** TOT.REM.   TOT.DED   TOT.NETO"
    Print #1, String(233, "-")
Next
Close #1

Call Imprime_Txt("REPORTELEGAL.txt", RUTA$)

End Sub
Private Sub cbotipobol_Click()
    If cbotipotrab.Text = "OBREROS" And cbotipobol.Text = "NORMAL" Then
        Label7.Visible = True: TxtSemana.Visible = True
    Else
        Label7.Visible = False: TxtSemana.Visible = False
    End If
End Sub
Private Sub cbotipotrab_Click()
    If cbotipotrab.Text = "OBREROS" And cbotipobol.Text = "NORMAL" Then
        Label7.Visible = True: TxtSemana.Visible = True
    Else
        Label7.Visible = False: TxtSemana.Visible = False
    End If
End Sub
Private Sub cmdprint_Click()
If Not optcab.Value Then
    Call Detalle_Resumen
Else
    Call Cabeza_Resumen
End If
End Sub

Private Sub optcab_Click()
    pctDet.Visible = True
    PctCab.Visible = False
End Sub
Private Sub optdetalle_Click()
    pctDet.Visible = False
    PctCab.Visible = True
End Sub
Private Sub Detalle_Resumen()

Dim scad As String
Dim scad2 As String
Dim scad3 As String
Dim scad4 As String
Dim Sql As String
Dim i As Long
Dim rd As ADODB.Recordset
Dim RUTA$
Dim filas As Integer
Dim i_Contador_Registros As Integer
Dim i_Dias_Trabajados As Integer
'Dim s_Tipo_Boleta As String
'Dim s_Tipo_Trabajador As String

i_Contador_Registros = 0
i_Contador_Registros_2 = 0

If txtaño = "" Then
    MsgBox "Debe Ingresar Año de Proceso", vbInformation
    txtaño.SetFocus
    Exit Sub
End If

If Val(txtaño) < 1950 Or Val(txtaño) > 2100 Then
    MsgBox "La informacion del año no es correcta", vbInformation
    txtaño.SetFocus
    Exit Sub
End If
'************************************************************************************

RUTA$ = App.path & "\REPORTS\" & "REPORTELEGALDET.txt"

If Not optcab.Value Then
    Sql = "SELECT p.placod,LTRIM(RTRIM(p.ap_pat))+' '+LTRIM(RTRIM(p.ap_mat))+' '+LTRIM(RTRIM(p.ap_cas))+' '+LTRIM(RTRIM(p.nom_1))+' '+LTRIM(RTRIM(p.nom_2)),m.descrip,p.fingreso,p.fcese,p.ipss,prb.importe,"
    Sql = Sql & " SUM(h01),SUM(h02),SUM(h03),SUM(h10),SUM(h11),0,0,SUM(h12),0,SUM(h17),"
    Sql = Sql & " SUM(i01),SUM(i09),SUM(i10),SUM(i02),SUM(i12),SUM(i07),SUM(i14),SUM(i15),SUM(i13),SUM(i11),0,0,SUM(i06+i05),SUM(i04),SUM(i29),0,0,0,0,0,SUM(i21),"
    Sql = Sql & " SUM(I03),0,SUM(I17),SUM(I18),SUM(I20),0,"
    Sql = Sql & " SUM(d01),SUM(d04),SUM(d02),SUM(d03),0,SUM(d09),SUM(d05),SUM(d07),SUM(d12),SUM(d13),SUM(d11),SUM(d06),SUM(a01),0,0,0,0,"
    Sql = Sql & " Sum(totaling), Sum(totalded), Sum(totneto) FROM PLAHISTORICO ph INNER JOIN PLANILLAS p ON ( p.cia=ph.cia AND  p.placod=ph.placod AND P.STATUS!='*' and tipotrabajador='" & Format(cbotipotrab.ItemData(cbotipotrab.ListIndex), "00") & "')"
    Sql = Sql & " INNER JOIN PLAREMUNBASE prb ON (prb.cia=p.cia and prb.placod=p.placod and prb.concepto='01' and prb.status!='*') INNER JOIN MAESTROS_3 m ON "
    Sql = Sql & " (m.ciamaestro='" & wcia & "055' AND m.status!='*' and m.cod_maestro3=p.cargo) WHERE PH.CIA='" & wcia & "' AND PH.STATUS!='*' AND MONTH(FECHAPROCESO)=" & cbomes.ItemData(cbomes.ListIndex)
    Sql = Sql & " AND YEAR(FECHAPROCESO)=" & txtaño.Text & " and ph.proceso='" & Format(cbotipobol.ItemData(cbotipobol.ListIndex), "00") & "' "
    If TxtSemana.Visible = True Then
        Sql = Sql & " and ph.semana='" & TxtSemana & "'"
    End If
    Sql = Sql & "GROUP BY p.placod,p.ap_pat,p.ap_mat,p.ap_cas,p.nom_1,p.nom_2,p.fingreso,p.fcese,p.ipss,p.cargo,prb.importe,m.descrip"
End If

Set rd = cn.Execute(Sql)
Dim XCONTADOR As Integer
XCONTADOR = 0
If Not rd.EOF Then
    filas = 0
    Open RUTA$ For Output As #1
    Call Imprimir_Titulo(True)

        Do While Not rd.EOF

            If i_Contador_Registros = 7 Then
                i_Contador_Registros = 0
                Print #1, ""
                Print #1, ""
                'Print #1, ""
                Call Imprimir_Titulo(True)
            End If

            '********codigo agregado giovanni 18092007**********************
            scad = "": scad2 = "  ": scad3 = " ": scad4 = " "
            Select Case cbotipobol.Text
                Case "NORMAL": s_Tipo_Boleta = "01"
                Case "VACACIONES": s_Tipo_Boleta = "02"
                Case "GRATIFICACION": s_Tipo_Boleta = "03"
            End Select
            Select Case cbotipotrab.Text
                Case "EMPLEADOS": s_Tipo_Trabajador = "01"
                Case "OBREROS": s_Tipo_Trabajador = "02"
            End Select
            Call Recupera_Dias_Trabajados(wcia, rd(0), cbomes.ListIndex + 1, _
            txtaño.Text, s_Tipo_Boleta, TxtSemana, s_Tipo_Trabajador)
            i_Dias_Trabajados = Reportes_Centrales.rs_RptCentrales_pub!dias
            Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
            '***************************************************************

            '*********codigo modificado giovanni 19092007*******************
            Print #1, ""

            For i = 0 To rd.Fields.count - 1
                Select Case i
                Case Is = 0:
                '   lentexto(65, Left(mtexto, 65))
                    scad = scad & Trim(rd(i)) & Space(8 - Len(Trim(rd(i))))
                    scad = lentexto(8, scad)
                Case Is = 1:
                    scad = scad & lentexto(40, Trim(rd(i)) & Space(40 - Len(Trim(rd(i)))))
                    'scad = lentexto(40, scad)
                Case Is = 2:
                    If Len(Trim(rd(i))) <= 25 Then
                        If Len(Trim(rd(i))) < 20 Then
                            'scad = scad & Trim(rd(i)) & Space(20 - Len(Trim(rd(i)))) & Space(5)
                            scad = scad & lentexto(30, Trim(rd(i)) & Space(30 - Len(Trim(rd(i)))) & Space(5))
                        Else
                            scad = scad & lentexto(30, Trim(rd(i)) & Space(30 - Len(Trim(rd(i)))) & Space(5))
                            
                        End If
                    Else
                        scad = scad & lentexto(30, Mid(rd(i), 1, 30) & Space(5))
                    End If
                Case 3, 4:
                    scad = scad & lentexto(10, Trim(rd(i) & "") & Space(10 - Len(Trim(rd(i) & ""))) & Space(2))
                Case Is = 5:
                    scad = scad & lentexto(21, Trim(rd(i)) & Space(21 - Len(Trim(rd(i)))))
                Case Is = 6:
                    Select Case cbotipotrab.Text
                        Case "EMPLEADOS"
                            scad = scad & Space(11)
                        Case "OBREROS"
                            scad = scad & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End Select
                    'scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(I)), "###,###,###.00")))
                Case 7
                    If Val(rd(i)) = 0 Then
                        scad = scad & Space(13)
                    Else
                        scad = scad & lentexto(13, Format(Trim(rd(i)), "###,###,###.00") & Space(13 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 8 To 16:
                    If Val(rd(i)) = 0 Then
                        scad = scad & Space(10)
                    Else
                        scad = scad & lentexto(10, Format(Trim(rd(i)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 17 To 22:
                    If Val(rd(i)) = 0 Then
                        scad2 = scad2 & Space(11)
                    Else
                        scad2 = scad2 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 23 To 28
                    If Val(rd(i)) = 0 Then
                        scad2 = scad2 & Space(12)
                    Else
                        scad2 = scad2 & lentexto(12, Format(Trim(rd(i)), "###,###,###.00") & Space(12 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 29 To 36
                    If Val(rd(i)) = 0 Then
                        scad2 = scad2 & Space(10)
                    Else
                        scad2 = scad2 & lentexto(10, Format(Trim(rd(i)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 37
                    If Val(rd(i)) = 0 Then
                        scad2 = scad2
                    Else
                        scad2 = scad2 & lentexto(6, Space(2) & Format(Trim(rd(i)), "###,###,###.00")) '& Space(Len(Format(Trim(rd(I)), "###,###,###.00")))
                    End If
                Case 38 To 42
                    If i = 42 Then
                        If Val(rd(i)) > 0 Then
                            scad3 = scad3 & lentexto(6, Format(Trim(rd(i)), "###,###,###.00") & Space(6 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                        Else
                            scad3 = scad3 & Space(6)
                        End If
                    Else
                        If Val(rd(i)) = 0 Then
                            scad3 = scad3 & Space(11)
                        Else
                            scad3 = scad3 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                        End If
                    End If
                Case 43
                    Select Case i_Dias_Trabajados
                        Case 0
                            scad3 = scad3 & Space(11)
                        Case Is < 31
                            scad3 = scad3 & lentexto(11, Space(1) & i_Dias_Trabajados & Space(11 - Len(i_Dias_Trabajados)))
                        Case Else
                            i_Dias_Trabajados = 30
                            scad3 = scad3 & lentexto(11, Space(1) & i_Dias_Trabajados & Space(11 - Len(i_Dias_Trabajados)))
                    End Select
                Case 44 To 45
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(11)
                    Else
                        scad4 = scad4 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 46 To 48
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(11)
                    Else
                        scad4 = scad4 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 49 To 51
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(10)
                    Else
                        scad4 = scad4 & lentexto(10, Format(Trim(rd(i)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 52
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(10)
                    Else
                        scad4 = scad4 & lentexto(10, Format(Trim(rd(i)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 53
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(12)
                    Else
                        scad4 = scad4 & lentexto(12, Format(Trim(rd(i)), "###,###,###.00") & Space(12 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 54
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(15)
                    Else
                        scad4 = scad4 & lentexto(15, Format(Trim(rd(i)), "###,###,###.00") & Space(15 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 55
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(18)
                    Else
                        scad4 = scad4 & lentexto(18, Format(Trim(rd(i)), "###,###,###.00") & Space(18 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 56 To 59
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(11)
                    Else
                        scad4 = scad4 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 60
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(17)
                    Else
                        scad4 = scad4 & lentexto(17, Format(Trim(rd(i)), "###,###,###.00") & Space(17 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 61 To 62
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(11)
                    Else
                        scad4 = scad4 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00") & Space(11 - Len(Format(Trim(rd(i)), "###,###,###.00"))))
                    End If
                Case 63
                    If Val(rd(i)) = 0 Then
                        scad4 = scad4 & Space(2)
                    Else
                        scad4 = scad4 & lentexto(11, Format(Trim(rd(i)), "###,###,###.00")) '& Space(Len(Format(Trim(rd(I)), "###,###,###.00")))
                    End If
                End Select
            Next i

            filas = filas + 1
            Print #1, scad
            Print #1, scad2
            Print #1, scad3
            Print #1, scad4
            Print #1, ""

            rd.MoveNext
            '*******************************************************************

            '**********codigo agregado giovanni 18092007******************
            i_Contador_Registros = i_Contador_Registros + 1
            i_Contador_Registros_2 = i_Contador_Registros_2 + 1

''            If i_Contador_Registros = 8 Then
''                i_Contador_Registros = 0
''                Print #1, ""
''                Print #1, ""
''                Print #1, ""
''                Call Imprimir_Titulo
''            End If
            '*************************************************************
If rd.EOF Then
    XCONTADOR = i_Contador_Registros
End If

        Loop

        Select Case XCONTADOR
            Case 1
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
               Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 2
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 3
                Print #1, "": Print #1, "": Print #1, ""
               Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 4
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 5
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
               Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 6
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 7
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
            Case 8
                Print #1, "": Print #1, "": Print #1, ""
        End Select

       Print #1, ""
       Print #1, ""
       Print #1, ""
       
        Call Imprimir_Titulo(False)
        Call Resumen_Planilla

    Close #1
    rd.Close
End If
Set rd = Nothing

Call Imprime_Txt("REPORTELEGALDET.txt", RUTA$)

End Sub
Sub Imprimir_Titulo(Optional PRINTCAB As Boolean = False)
    Dim i_Contador_Titulo As Integer
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Select Case cbotipotrab.Text
        Case "EMPLEADOS"
     '   lentexto(65, Left(mtexto, 65))
            Print #1, Space(70) & "Mes de " & cbomes.Text & " del " & txtaño
           
          ' Print #1, lentexto(70, Left("Mes de " & cbomes.Text & " del " & txtaño, 70))
           
        Case "OBREROS"
            If cbotipobol.ItemData(cbotipobol.ListIndex) = "1" Then
                'Print #1, Space(70) & "Mes de " & cbomes.Text & " del " & TxtAño & " Semana N " & TxtSemana
                Print #1, Space(70) & lentexto(70, Left("Mes de " & cbomes.Text & " del " & txtaño & " Semana N " & TxtSemana, 70))
            Else
                'Print #1, Space(70) & "Mes de " & cbomes.Text & " del " & TxtAño
                Print #1, Space(70) & lentexto(70, "Mes de " & cbomes.Text & " del " & txtaño)
            End If
    End Select
'    Print #1, Space(70) & "Mes de " & cbomes.Text & " del " & txtaño
    Print #1, ""
    For i_Contador_Titulo = 0 To 8
        If i_Contador_Titulo <> 5 Then
            Print #1, ""
        Else
            If PRINTCAB Then
                Print #1, Space(42) & "CONTRB"
            Else
                Print #1, ""
            End If
        End If
    Next
End Sub
Private Sub Resumen_Planilla()
Dim rs2 As ADODB.Recordset
Dim rsremu As ADODB.Recordset
Dim s_sql_pla As String
Dim rscargo As ADODB.Recordset
Dim mmes As Integer
Dim mano As Integer
Dim msem As String
Dim i As Integer
Dim mcad As String
Dim MCADIT As String
Dim mcadI As String
Dim mcadd As String
Dim mcaddt As String
Dim mcada As String
Dim mcadh As String
Dim mcadat As String
Dim mcadht As String
Dim mfor As Integer
Dim mc As Integer
Dim mcargo As String
Dim mremun As Currency
Dim numtra As Integer
Dim rs As New ADODB.Recordset
Dim Inicio As Boolean

'****************************************************************************

Dim mcadir As String
Dim mcadinr As String
Dim cad As String
Dim Rcadh As String
Dim Vcadh As String
Dim Rcadir As String
Dim Vcadir As String
Dim Rcadinr As String
Dim Vcadinr As String
Dim Rcadd As String
Dim Vcadd As String
Dim Rcada As String
Dim Vcada As String

Dim s_mcadh As String
Dim s_mcadI As String
Dim s_mcadd As String
Dim s_mcada As String

mcadir = "": mcadd = "": mcada = "": mcadinr = "": mcadh = ""
Vcadir = ""
Vcadd = ""
Vcada = ""
Vcadinr = ""
Vcadh = ""
Rcadir = ""
Rcadd = ""
Rcada = ""
Rcadinr = ""
Rcadh = ""

Dim MCADENA As String

If cbotipobol.Text = "TOTAL" Then
    MCADENA = ""
Else
    MCADENA = " AND CHARINDEX('" & Left(cbotipobol.Text, 1) & "',flag1 )>0 "
End If

s_sql_pla = "select a.codigo,b.descrip from plaverhoras a," & _
       "maestros_2 b where b.ciamaestro='01077' " & _
       "and a.cia='" & wcia & "' and a.tipo_trab='" & _
       s_Tipo_Trabajador & "' and a.status<>'*' " & MCADENA _
       & "and a.codigo=b.cod_maestro2 order by codigo"
       
If (fAbrRst(rs2, s_sql_pla)) Then rs2.MoveFirst

'If s_Tipo_Boleta = "03" Then
'       mcadh = ""
'       Rcadh = ""
'       Vcadh = ""
'Else
    Do While Not rs2.EOF
       mcadh = mcadh & " " & lentexto(9, Left(rs2!descrip, 9)) & " "
       Rcadh = Rcadh & lentexto(20, Left(rs2!descrip, 20))
       Vcadh = Vcadh & rs2!codigo
       rs2.MoveNext
    Loop
'End If

If rs2.State = 1 Then rs2.Close

'INGRESOS
s_sql_pla = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
     & "where a.cia='" & wcia & "' and a.tipo='I' and a.status<>'*' " _
     & "and a.cia=b.cia and b.status<>'*' " & _
     "and b.tipomovimiento='02' and " & _
     "a.tipo_trab='" & s_Tipo_Trabajador & "' and a.codigo=b.codinterno " _
     & "order by a.tipo,a.codigo"

If (fAbrRst(rs2, s_sql_pla)) Then rs2.MoveFirst
Do While Not rs2.EOF
   s_sql_pla = "select cod_remu from plaafectos where cia='" & _
   wcia & "' and status<>'*' and cod_remu='" & rs2!codigo & _
   "' and tipo in ('A','D') AND CODIGO!='13' "
   If (fAbrRst(rsremu, s_sql_pla)) Then
      mcadir = mcadir & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Rcadir = Rcadir & lentexto(20, Left(rs2!Descripcion, 20))
      Vcadir = Vcadir & rs2!codigo
   Else
      Rcadinr = Rcadinr & lentexto(20, Left(rs2!Descripcion, 20))
      mcadinr = mcadinr & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Vcadinr = Vcadinr & rs2!codigo
   End If
   If rsremu.State = 1 Then rsremu.Close
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

'DEDUCCIONES Y APORTACIONES
s_sql_pla = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
     & "where a.cia='" & wcia & "' and a.tipo<>'I' and a.status<>'*' " _
     & "and a.cia=b.cia and b.status<>'*' and b.tipomovimiento='03' and a.tipo_trab='" & s_Tipo_Trabajador & "' and a.codigo=b.codinterno " _
     & "order by a.tipo,a.codigo"
    
If (fAbrRst(rs2, s_sql_pla)) Then rs2.MoveFirst

Do While Not rs2.EOF
   Select Case rs2!tipo
          Case Is = "D"
               mcadd = mcadd & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Rcadd = Rcadd & lentexto(20, Left(rs2!Descripcion, 20))
               Vcadd = Vcadd & rs2!codigo
          Case Is = "A"
               Rcada = Rcada & lentexto(20, Left(rs2!Descripcion, 20))
               mcada = mcada & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Vcada = Vcada & rs2!codigo
   End Select
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

'****************************************************************************

For i = 1 To 50
   mcadI = mcadI & "sum(i" & Format(i, "00") & ") as i" & Format(i, "00") & ","
   MCADIT = MCADIT & "i" & Format(i, "00") & "+"
   If i <= 30 Then
      If i <= 20 Then
         'mcadd = mcadd & "sum(d" & Format(I, "00") & ") as d" & Format(I, "00") & ","
         'mcada = mcada & "sum(a" & Format(I, "00") & ") as a" & Format(I, "00") & ","
         'mcaddt = mcaddt & "d" & Format(I, "00") & "+"
         'mcadat = mcadat & "a" & Format(I, "00") & "+"
         s_mcadd = s_mcadd & "sum(d" & Format(i, "00") & ") as d" & Format(i, "00") & ","
         s_mcada = s_mcada & "sum(a" & Format(i, "00") & ") as a" & Format(i, "00") & ","
         mcaddt = mcaddt & "d" & Format(i, "00") & "+"
         mcadat = mcadat & "a" & Format(i, "00") & "+"
      End If
      s_mcadh = s_mcadh & "sum(h" & Format(i, "00") & ") as h" & Format(i, "00") & ","
      mcadht = mcadht & "h" & Format(i, "00") & "+"
   End If
Next

s_mcadI = Mid(mcadI, 1, Len(Trim(mcadI)) - 1)
s_mcadd = Mid(s_mcadd, 1, Len(Trim(s_mcadd)) - 1)
s_mcada = Mid(s_mcada, 1, Len(Trim(s_mcada)) - 1)
s_mcadh = Mid(s_mcadh, 1, Len(Trim(s_mcadh)) - 1)

Print #1,
Print #1, Space(10) & "***** TOTAL PLANILLA *****"
Print #1,
Print #1, "                           H O R A S                             R E M U N E R A C I O N E S                        D E D U C C I O N E S                           A P O R T A C I O N E S"
Print #1, "                           ---------                             ---------------------------                        ---------------------                           -----------------------"
If rs.State = 1 Then rs.Close
         
s_sql_pla = ""
s_sql_pla = "select " & s_mcadh & "," & s_mcadI & "," & s_mcadd & "," & s_mcada
s_sql_pla = s_sql_pla & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto "
s_sql_pla = s_sql_pla & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & txtaño & " and month(fechaproceso)=" & cbomes.ListIndex + 1 & " and proceso LIKE '" & s_Tipo_Boleta + "%" & "' and semana LIKE '" & TxtSemana + "%" & "' " _
& "and tipotrab='" & s_Tipo_Trabajador & "' and status<>'*'"

If (fAbrRst(rs, s_sql_pla)) Then

Dim m As Integer
m = 0
mremun = 0

For i = 1 To 100 Step 2
   m = m + 1
   
   'TOTAL HORAS
   mcad = Space(10)
   If i <= Len(Vcadh) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadh, 1, 20)
      Else
         mcad = mcad & Mid(Rcadh, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadh, i, 2))
      If rs(mc - 1) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc - 1), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(39)
   End If
   
   '=========================
   
   'TOTAL REMUNERACIONES AFECTAS
   mcad = mcad & Space(10)
   If i <= Len(Vcadir) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadir, 1, 20)
      Else
         mcad = mcad & Mid(Rcadir, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadir, i, 2))
      If rs(mc + 29) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 29), "###,###,##0.00")
         mremun = mremun + IIf(IsNull(rs(mc + 29)), 0, rs(mc + 29))
      End If
   Else
      mcad = mcad & Space(14)
   End If
   
   'TOTAL DEDUCCIONES
   mcad = mcad & Space(10)
   If i <= Len(Vcadd) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadd, 1, 20)
      Else
         mcad = mcad & Mid(Rcadd, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadd, i, 2))
      If rs(mc + 79) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 79), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(14)
   End If
   
   'TOTAL APORTACIONES
   mcad = mcad & Space(10)
   If i <= Len(Vcadir) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcada, 1, 20)
      Else
         mcad = mcad & Mid(Rcada, 20 * (m - 1) + 1, 20)
      End If
      mc = Val(Mid(Vcada, i, 2))
      If rs(mc + 99) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 99), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(14)
   End If

   If Trim(mcad) <> "" Then Print #1, mcad
'   Debug.Print mcad
Next
Print #1, Space(84) & "--------------"
Print #1, Space(59) & "Sub Total I          " & fCadNum(mremun, "###,###,###,##0.00")
Print #1, Space(65) & "NO REMUNERATIVOS"
Print #1, Space(65) & "----------------"
mremun = 0


For i = 1 To Len(Vcadinr) Step 2
   'TOTAL REMUNERACIONES NO AFECTAS
   mcad = ""
   mcad = mcad & Space(59)
   If i <= Len(Vcadinr) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadinr, 1, 20)
      Else
         mcad = mcad & Mid(Rcadinr, 20 * (i - 2) + 1, 20)
      End If
      
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadinr, i, 2))
      If rs(mc + 29) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(IIf(IsNull(rs(mc + 29)), 0, rs(mc + 29)), "###,###,##0.00")
         mremun = mremun + IIf(IsNull(rs(mc + 29)), 0, rs(mc + 29))
      End If
   Else
      mcad = mcad & Space(14)
   End If
   If Trim(mcad) <> "" Then Print #1, mcad
Next i

Print #1, Space(84) & "--------------"
Print #1, Space(59) & "Sub Total II         " & fCadNum(mremun, "###,###,###,##0.00")
Print #1, Space(84) & "--------------                                   --------------                                 -----------"
Print #1, Space(59) & "* TOTAL REMUNERACION *" & fCadNum(rs!totaling, "##,###,###,##0.00") & Space(10) & "* TOTAL DEDUCCIONES * " & fCadNum(rs!totalded, "##,###,###,##0.00") & Space(10) & "* TOTAL APORTACIONES *" & fCadNum(rs!totalapo, "#,###,##0.00")
Print #1,
Print #1, Space(59) & "*** NETO PAGADO ***   " & fCadNum(rs!totneto, "##,###,###,##0.00")
Print #1, Space(59) & "======================================="
Print #1,
Print #1, "Total de Trabajadores    =>  " & Str(i_Contador_Registros_2)
Print #1, Chr(12) + Chr(13)
End If

End Sub


'Private Sub Detalle_Resumen()
'
'Dim scad As String
'Dim scad2 As String
'Dim scad3 As String
'Dim scad4 As String
'Dim sql As String
'Dim I As Long
'Dim rd As ADODB.Recordset
'Dim RUTA$
'Dim filas As Integer
'Dim i_Contador_Registros As Integer
'Dim i_Dias_Trabajados As Integer
''Dim s_Tipo_Boleta As String
''Dim s_Tipo_Trabajador As String
'
'i_Contador_Registros = 0
'i_Contador_Registros_2 = 0
'
'If TxtAño = "" Then
'    MsgBox "Debe Ingresar Año de Proceso", vbInformation
'    TxtAño.SetFocus
'    Exit Sub
'End If
'
'If Val(TxtAño) < 1950 Or Val(TxtAño) > 2100 Then
'    MsgBox "La informacion del año no es correcta", vbInformation
'    TxtAño.SetFocus
'    Exit Sub
'End If
''************************************************************************************
'
'RUTA$ = App.Path & "\REPORTS\" & "REPORTELEGALDET.txt"
'
'If Not optcab.Value Then
'    sql = "SELECT p.placod,LTRIM(RTRIM(p.ap_pat))+' '+LTRIM(RTRIM(p.ap_mat))+' '+LTRIM(RTRIM(p.ap_cas))+' '+LTRIM(RTRIM(p.nom_1))+' '+LTRIM(RTRIM(p.nom_2)),m.descrip,p.fingreso,p.fcese,p.ipss,prb.importe,"
'    sql = sql & " SUM(h01),SUM(h02),SUM(h03),SUM(h10),SUM(h11),0,0,SUM(h12),0,SUM(h17),"
'    sql = sql & " SUM(i01),SUM(i09),SUM(i10),SUM(i02),SUM(i12),SUM(i07),SUM(i14),SUM(i15),SUM(i13),SUM(i11),0,0,SUM(i06+i05),SUM(i04),SUM(i29),0,0,0,0,0,SUM(i21),"
'    sql = sql & " SUM(I03),0,SUM(I17),SUM(I18),0,"
'    sql = sql & " SUM(d01),SUM(d04),SUM(d02),SUM(d03),0,SUM(d09),SUM(d05),SUM(d07),SUM(d12),SUM(d13),SUM(d11),SUM(d06),SUM(a01),0,0,0,0,"
'    sql = sql & " Sum(totaling), Sum(totalded), Sum(totneto) FROM PLAHISTORICO ph INNER JOIN PLANILLAS p ON ( p.cia=ph.cia AND  p.placod=ph.placod AND P.STATUS!='*' and tipotrabajador='" & Format(cbotipotrab.ItemData(cbotipotrab.ListIndex), "00") & "')"
'    sql = sql & " INNER JOIN PLAREMUNBASE prb ON (prb.cia=p.cia and prb.placod=p.placod and prb.concepto='01' and prb.status!='*') INNER JOIN MAESTROS_3 m ON "
'    sql = sql & " (m.ciamaestro='" & wcia & "055' AND m.status!='*' and m.cod_maestro3=p.cargo) WHERE PH.CIA='" & wcia & "' AND PH.STATUS!='*' AND MONTH(FECHAPROCESO)=" & cbomes.ItemData(cbomes.ListIndex)
'    sql = sql & " AND YEAR(FECHAPROCESO)=" & TxtAño.Text & " and ph.proceso='" & Format(cbotipobol.ItemData(cbotipobol.ListIndex), "00") & "' "
'    If TxtSemana.Visible = True Then
'        sql = sql & " and ph.semana='" & TxtSemana & "'"
'    End If
'    sql = sql & "GROUP BY p.placod,p.ap_pat,p.ap_mat,p.ap_cas,p.nom_1,p.nom_2,p.fingreso,p.fcese,p.ipss,p.cargo,prb.importe,m.descrip"
'End If
'
'Set rd = cn.Execute(sql)
'
'If Not rd.EOF Then
'        filas = 0
'    Open RUTA$ For Output As #1
'    For I = 0 To 15
'        Print #1, ""
'    Next
'        Do While Not rd.EOF
'            scad = ""
'            If filas = 9 Then
'                Print #1, Chr(12) + Chr(13)
'                For I = 0 To 15
'                    Print #1, ""
'                Next
'                filas = 0
'            End If
'            For I = 0 To rd.Fields.count - 1
'                Select Case I
'                Case Is = 0:
'                    scad = scad & Trim(rd(I)) & Space(8 - Len(Trim(rd(I))))
'                Case Is = 1:
'                    scad = scad & Trim(rd(I)) & Space(40 - Len(Trim(rd(I))))
'                Case Is = 2:
'                    If Len(Trim(rd(I))) <= 25 Then
'                        scad = scad & Trim(rd(I)) & Space(20 - Len(Trim(rd(I)))) & Space(5)
'                    Else
'                        scad = scad & Mid(rd(I), 1, 20) & Space(5)
'                    End If
'                Case 3, 4:
'                    scad = scad & Trim(rd(I) & "") & Space(11 - Len(Trim(rd(I) & ""))) & Space(2)
'                Case Is = 5:
'                    scad = scad & Trim(rd(I)) & Space(19 - Len(Trim(rd(I))))
'                Case Is = 6:
'                    scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(I)), "###,###,###.00")))
'                Case 7 To 16:
'                    scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(I)), "###,###,###.00")))
'                    If I = 16 Then scad = scad & vbCrLf
'                Case 17 To 37:
'                    scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(I)), "###,###,###.00")))
'                    If I = 37 Then scad = scad & vbCrLf
'                Case 38 To 42
'                    scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(I)), "###,###,###.00")))
'                    If I = 42 Then scad = scad & vbCrLf
'                Case 43 To 58
'                    scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(I)), "###,###,###.00")))
'                Case 59 To 61
'                    scad = scad & Format(Trim(rd(I)), "###,###,###.00") & Space(10 - Len(Format(Trim(rd(I)), "###,###,###.00")))
'                End Select
'            Next I
'            filas = filas + 1
'            Print #1, scad
'            Print #1, ""
'            rd.MoveNext
'        Loop
'    Close #1
'    rd.Close
'End If
'Set rd = Nothing
'
'
''        Print #1, ""
' '       Print #1, ""
''  '      Print #1, ""
''        Call Imprimir_Titulo
''        Call Resumen_Planilla
'
''    Close #1
''    rd.Close
''Set rd = Nothing
'
'Call Imprime_Txt("REPORTELEGALDET.txt", RUTA$)
'
'End Sub
'
