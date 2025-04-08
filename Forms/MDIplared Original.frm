VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIplared 
   BackColor       =   &H8000000C&
   Caption         =   "» Sistema de Planilla"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Enabled         =   0   'False
   Icon            =   "MDIplared.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList Img_00 
      Left            =   90
      Top             =   555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":0E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":13D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":1972
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":1F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":24A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIplared.frx":2A40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7050
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "Usuario"
            TextSave        =   "Usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "Servidor"
            TextSave        =   "Servidor"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "BD"
            TextSave        =   "BD"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Empresa"
            TextSave        =   "Empresa"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      ButtonWidth     =   2090
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "Img_00"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo  "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar  "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar  "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar  "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir  "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesar  "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir  "
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compañias"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Carga Derechohabientes"
         Height          =   375
         Left            =   10560
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Menu mnumaestro 
      Caption         =   "&Maestros"
      Tag             =   "M1"
      Begin VB.Menu Mnupersona 
         Caption         =   "&Personal"
      End
      Begin VB.Menu MnuConder 
         Caption         =   "&Consulta DerechoHabientes"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconcal 
         Caption         =   "&Deducciones y Aportaciones"
      End
      Begin VB.Menu mnuconremun 
         Caption         =   "Conceptos &Remunerativos"
      End
      Begin VB.Menu mnuhorpla 
         Caption         =   "Ho&ras de Planilla"
      End
      Begin VB.Menu spcia 
         Caption         =   "-"
      End
      Begin VB.Menu mnucia 
         Caption         =   "&Compañias"
      End
      Begin VB.Menu mnuobras 
         Caption         =   "&Obras"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhoraper 
         Caption         =   "&Horas por Periodo"
      End
      Begin VB.Menu mnuperpago 
         Caption         =   "&Periodos de Pago"
      End
      Begin VB.Menu mnumae 
         Caption         =   "Tablas Maestras"
      End
      Begin VB.Menu mnuconceptos 
         Caption         =   "Conceptos Generales"
      End
      Begin VB.Menu mnufacsenati 
         Caption         =   "Factores Para Senati"
         Visible         =   0   'False
      End
      Begin VB.Menu mnutc 
         Caption         =   "Tipo de Cambio"
         Visible         =   0   'False
      End
      Begin VB.Menu mnucargo 
         Caption         =   "Cargo Personal"
      End
      Begin VB.Menu spm1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuusers 
         Caption         =   "Administar Usuarios"
      End
   End
   Begin VB.Menu mnuformpara 
      Caption         =   "&Parametros"
      Tag             =   "M1"
      Begin VB.Menu mnuafp 
         Caption         =   "&AFP"
      End
      Begin VB.Menu mnusctr 
         Caption         =   "SCTR-Seg. Com. Trab. Riesgo"
      End
      Begin VB.Menu mnutasascrt 
         Caption         =   "&Tasa de Calculo de SCRT"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuuit 
         Caption         =   "&UIT - Sueldo Minimo"
      End
      Begin VB.Menu mnufactor 
         Caption         =   "&Factor de Calculo"
      End
      Begin VB.Menu mnuprom 
         Caption         =   "&Calculo de Promedios (Vacaciones)"
      End
      Begin VB.Menu mnueps 
         Caption         =   "Entidad Prestadora de Salud - E.P.S"
      End
      Begin VB.Menu spformula 
         Caption         =   "-"
      End
      Begin VB.Menu mnuafectas 
         Caption         =   "Remuneraciones Afectas"
      End
      Begin VB.Menu mnuforremu 
         Caption         =   "&Formulas de Conceptos Remunerativos"
         Visible         =   0   'False
      End
      Begin VB.Menu sphorcia 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhorcia 
         Caption         =   "Horas por Compañia"
      End
      Begin VB.Menu spsem 
         Caption         =   "-"
      End
      Begin VB.Menu mnuiniciosem 
         Caption         =   "Inicio Anual de Semana / Mes"
      End
      Begin VB.Menu spprint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuseteobol 
         Caption         =   "Seteo de Impresion de Boletas"
      End
      Begin VB.Menu spcts 
         Caption         =   "-"
      End
      Begin VB.Menu mnucts 
         Caption         =   "CTS"
      End
      Begin VB.Menu spferiados 
         Caption         =   "-"
      End
      Begin VB.Menu mnuregfer 
         Caption         =   "Registro de Feriados"
      End
   End
   Begin VB.Menu mnumov 
      Caption         =   "Mo&vimientos"
      Tag             =   "M1"
      Begin VB.Menu mnucalcbol 
         Caption         =   "&Calculo de Boletas"
      End
      Begin VB.Menu mnuprintbol 
         Caption         =   "Impresion de Boletas"
      End
      Begin VB.Menu mn_crystal 
         Caption         =   "Impresion Crystal"
         Visible         =   0   'False
      End
      Begin VB.Menu spctacte0 
         Caption         =   "-"
      End
      Begin VB.Menu mnupresta 
         Caption         =   "&Prestamos al trabajador"
      End
      Begin VB.Menu spmov1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquincena 
         Caption         =   "Adelanto de Quincena"
      End
      Begin VB.Menu mnuprintquinc 
         Caption         =   "Impresion de Quincenas"
      End
      Begin VB.Menu spprintbol 
         Caption         =   "-"
      End
      Begin VB.Menu mnutareo 
         Caption         =   "Tareo Diario"
      End
      Begin VB.Menu spotras 
         Caption         =   "-"
      End
      Begin VB.Menu mnubolmas 
         Caption         =   "Generaion de Boletas Masivas"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusub 
         Caption         =   "Subsidios y Liquidaciones"
      End
      Begin VB.Menu MnuDiasSub 
         Caption         =   "Dias Subsidiados del trabajador"
      End
      Begin VB.Menu MnuDiasNoTrab 
         Caption         =   "Dias no Trabajados y no subsidiados del trabajador"
      End
      Begin VB.Menu mnudepo 
         Caption         =   "Archivo Depositos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnucont 
         Caption         =   "Emision de Contrato"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubilletaje 
         Caption         =   "Billetaje"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuconsul 
      Caption         =   "Consultas y &Reportes"
      Begin VB.Menu mnuconsultareo 
         Caption         =   "&Tareo"
      End
      Begin VB.Menu mnuafprep 
         Caption         =   "Reportes de AFP"
      End
      Begin VB.Menu mnugenafpnet 
         Caption         =   "Reporte de AFP-Net"
      End
      Begin VB.Menu mnuanualafp 
         Caption         =   "Liquidacion Anual de Aportes de Afp"
      End
      Begin VB.Menu mnucalcsegvida 
         Caption         =   "Calculo de Seguro de Vida"
      End
      Begin VB.Menu mnuresumen 
         Caption         =   "Resumen de Planillas"
      End
      Begin VB.Menu mnulegal 
         Caption         =   "Planilla Legalizada"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudetquinta 
         Caption         =   "Detalle de Quinta Categoria"
      End
      Begin VB.Menu mnudeducapor 
         Caption         =   "Deducciones y Aportaciones Mensuales"
      End
      Begin VB.Menu mnuapordeducanual 
         Caption         =   "Deducciones y Aportaciones Anuales"
      End
      Begin VB.Menu mnuremun 
         Caption         =   "Remuneraciones"
      End
      Begin VB.Menu mnuselecc 
         Caption         =   "Consulta por Seleccion de Conceptos"
      End
      Begin VB.Menu mnuseg 
         Caption         =   "Seguro Riesgo Salud"
      End
      Begin VB.Menu Mnu_RptAsistencia 
         Caption         =   "Asistencias Personal"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ListTrabajadores 
         Caption         =   "Lista de Trabajadores"
      End
      Begin VB.Menu Mnu_AportesPrevGen 
         Caption         =   "Aportes Previsionales General"
      End
      Begin VB.Menu Mnu_DetPromedios 
         Caption         =   "Detalles de Promedios "
      End
      Begin VB.Menu mnuConceptoRemu 
         Caption         =   "Detalle de Conceptos Remunerativos"
      End
      Begin VB.Menu mnuremu_cts 
         Caption         =   "Remuneración - CTS"
      End
      Begin VB.Menu Mnu_CompRetenciones 
         Caption         =   "Reporte Comprobante Retenciones"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuhorasextras 
         Caption         =   "Horas Extras"
      End
      Begin VB.Menu mnuhistoricos 
         Caption         =   "Otros Reportes"
         Begin VB.Menu mnuplamas_plahistorico 
            Caption         =   "Plamas - Plahistorico"
         End
      End
      Begin VB.Menu SPC1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusunat 
         Caption         =   "Exportar a Sunat"
         Begin VB.Menu mnusunattrab 
            Caption         =   "Relacion de Trabajadores"
         End
         Begin VB.Menu Mnuderchohab 
            Caption         =   "Relacion de Derechohabientes"
         End
         Begin VB.Menu mnuremunrasunat 
            Caption         =   "Remuneraciones"
         End
         Begin VB.Menu MnuPlaElec 
            Caption         =   "Planilla Electronica"
         End
      End
      Begin VB.Menu mnu_03 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuestad 
         Caption         =   "Estadisticos"
         Visible         =   0   'False
         Begin VB.Menu mnuvf 
            Caption         =   "Resumen de Planillas (Cuadro IV F)"
         End
         Begin VB.Menu mnucudiv 
            Caption         =   "Horas Efectivas de Trabajo  (Cuadro IV)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu spcert 
         Caption         =   "-"
      End
      Begin VB.Menu mnucertifica 
         Caption         =   "Certificados"
         Begin VB.Menu mnucerqta 
            Caption         =   "Retencion de Quinta Categoria"
         End
         Begin VB.Menu mnucerafp 
            Caption         =   "Retencion de Sistema de Pensiones"
         End
      End
      Begin VB.Menu spctsr 
         Caption         =   "-"
      End
      Begin VB.Menu mnudepprov 
         Caption         =   "Depositos y Provisiones"
         Begin VB.Menu mnudepcts 
            Caption         =   "Provisión y Depositos de CTS"
         End
         Begin VB.Menu mnuprovac 
            Caption         =   "Provision de Vacaciones"
         End
         Begin VB.Menu mnuprovgrati 
            Caption         =   "Provision de Gratificacion"
         End
      End
      Begin VB.Menu spdev 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvacdev 
         Caption         =   "Vacaciones &Devengadas"
      End
      Begin VB.Menu Mnu_AnualEssalud 
         Caption         =   "Resumen Anual Essalud"
      End
      Begin VB.Menu CCC 
         Caption         =   "Remuneracion Acumulada"
      End
      Begin VB.Menu Mnu_Consulta1 
         Caption         =   "Consulta Promedio 1"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuCtsFalta 
         Caption         =   "Faltas dias CTS"
      End
   End
   Begin VB.Menu mnuventana 
      Caption         =   "Ve&ntana"
      Tag             =   "M1"
      WindowList      =   -1  'True
      Begin VB.Menu mnuhor 
         Caption         =   "Mosaico Horizontal"
      End
      Begin VB.Menu mnuver 
         Caption         =   "Mosaico Vertical"
      End
      Begin VB.Menu mnucasca 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mnuorgico 
         Caption         =   "Organizar Iconos"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuhelpcont 
         Caption         =   "Contenido"
      End
      Begin VB.Menu sphelp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuacer 
         Caption         =   "Acerca de Sistema de Planilla"
      End
   End
   Begin VB.Menu mnuproceso 
      Caption         =   "Procesos"
      Tag             =   "M1"
      Begin VB.Menu mnucarga 
         Caption         =   "Carga Datos"
      End
      Begin VB.Menu mnubackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnurestaura 
         Caption         =   "Restaurar Backup"
      End
      Begin VB.Menu mnucargapla 
         Caption         =   "Carga Planillas"
      End
   End
   Begin VB.Menu Mnu_Contabilidad 
      Caption         =   "Contabilidad"
      Begin VB.Menu Mnu_ContableMaest 
         Caption         =   "Ingreso Contable Maestro"
      End
      Begin VB.Menu Mnu_Parametros 
         Caption         =   "Parametros Adicionales"
         Begin VB.Menu Mnu_ContFijo 
            Caption         =   "Ingreso Contable Fijo"
         End
         Begin VB.Menu Mnu_Asociadas 
            Caption         =   "Cuentas Asociadas"
         End
      End
      Begin VB.Menu Mnu_Provisiones 
         Caption         =   "Ingreso Provisiones"
      End
      Begin VB.Menu mnucencos 
         Caption         =   "Centros de Costos"
      End
      Begin VB.Menu mnu_00 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Reportes 
         Caption         =   "Reporte Contables"
         Begin VB.Menu Mnu_RptDetallado 
            Caption         =   "Asientos Detallado"
            Enabled         =   0   'False
         End
         Begin VB.Menu Mnu_AsientosRptGeneral 
            Caption         =   "Asientos_General"
         End
      End
      Begin VB.Menu mnu_01 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_GenAsientos 
         Caption         =   "Generar Asientos Contables"
      End
      Begin VB.Menu Mnu_ArchivoDbf 
         Caption         =   "Generar Archivo DBF"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexportar 
         Caption         =   "Exportar Asientos"
      End
   End
End
Attribute VB_Name = "MDIplared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CCC_Click()
    FrmRemunAcumulada.Show
End Sub

Private Sub Command1_Click()
If UCase(Trim(wuser)) <> "SA" Then Exit Sub
'Actualiza_DatosTrabajador
'Actualiza_DerechoHabientes
Carga_DerechoHabientes
End Sub

Private Sub MDIForm_Load()
'MDIplared.Caption = MDIplared.Caption & Space(5) & UCase(gsSQL_Server) & " - " & UCase(gsSQL_DB)
'MDIplared.StatusBar1.Panels(2).Text = wuser
'MhState1.Value = True
    Me.Caption = "» Sistema de Planilla" & Space(40) & "Versión " & "[ " & App.Major & "." & App.Minor & "." & App.Revision & " ]"
    Call Setea_StatusBar

TitMsg = "Sistema de Planilla"

End Sub

Public Sub Setea_StatusBar()
    With StatusBar
        .Panels(1).Width = 800
        .Panels(2).Width = 1400
        .Panels(3).Width = 800
        .Panels(4).Width = 1400
        .Panels(5).Width = 500
        .Panels(6).Width = 2000
        .Panels(7).Width = 1000
        .Panels(8).Width = 500
        .Panels(9).Width = 4000
        
        .Panels(1).Text = "Usuario"
        .Panels(2).Text = UCase(wuser)
        .Panels(3).Text = "Servidor"
        .Panels(4).Text = UCase(gsSQL_Server)
        .Panels(5).Text = "BD"
        .Panels(6).Text = UCase(gsSQL_DB)
        .Panels(7).Text = "Empresa"
        .Panels(8).Text = wcia
        .Panels(9).Text = Trae_CIA(wcia)
    End With
End Sub

Private Sub mn_crystal_Click()
Unload frmprintcrystal
frmprintcrystal.wmBolQuin = "B"
frmprintcrystal.Show
frmprintcrystal.ZOrder 0
End Sub

Private Sub Mnu_AnualEssalud_Click()
    FrmEssaludAnual.Show
End Sub

Private Sub Mnu_AportesPrevGen_Click()
    FrmRptAportes.Show
End Sub
Private Sub Mnu_ArchivoDbf_Click()
    Call FrmGenerarDBF.Show
End Sub
Private Sub Mnu_AsientosRptGeneral_Click()
    FrmMGenera.Caption = "Imprimir Asientos - GENERAL "
    FrmMGenera.Show
End Sub
Private Sub Mnu_Asociadas_Click()
    FrmMContAsoc.Show
End Sub
Private Sub Mnu_CompRetenciones_Click()
    FrmRptRetenciones.Show
End Sub

Private Sub Mnu_Consulta1_Click()
'frm_Cons1.MDIChild = True
frm_Cons1.Show
End Sub

Private Sub Mnu_ContableMaest_Click()
    FrmMContable.Show
End Sub
Private Sub Mnu_ContFijo_Click()
    FrmMContableFijo.Show
End Sub
Private Sub Mnu_DetPromedios_Click()
    FrmPromedios.Show
End Sub
Private Sub Mnu_GenAsientos_Click()
    FrmMGenera.Caption = "Generar Asientos Contables"
    FrmMGenera.Show
End Sub
Private Sub Mnu_ListTrabajadores_Click()
    FrmRptseleccion2.Show
End Sub
Private Sub Mnu_Provisiones_Click()
    FrmCtaProvision.Show
End Sub
Private Sub Mnu_RptAsistencia_Click()
    FrmRptSeleccion.Show
End Sub
Private Sub Mnu_RptDetallado_Click()
    FrmMGenera.Caption = "Imprimir Asientos - DETALLADO"
    FrmMGenera.Show
End Sub
Private Sub mnuafectas_Click()
Load FrmAfectos
FrmAfectos.Show
FrmAfectos.ZOrder 0
End Sub
Private Sub mnuafp_Click()
Load FrmAfp
FrmAfp.Show
FrmAfp.ZOrder 0
End Sub
Private Sub mnuafprep_Click()
Load FrmRepafp
FrmRepafp.Show
FrmRepafp.ZOrder 0
End Sub

Private Sub mnuaporsenati_Click()
NameForm = "APORTESENATI"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnuanualafp_Click()
Load Frmafpaunal
Frmafpaunal.Show
Frmafpaunal.ZOrder 0
End Sub

Private Sub mnuapordeducanual_Click()
NameForm = "DEDUCAPORANUAL"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnubackup_Click()
Unload Frmbackup
Frmbackup.Opc = "B"
Load Frmbackup
Frmbackup.Show
Frmbackup.ZOrder 0
End Sub

Private Sub mnubilletaje_Click()
Unload frmbillete
frmbillete.Show
End Sub

Private Sub mnubolmas_Click()
Unload frmboletamasiva
frmboletamasiva.Show
End Sub

Private Sub mnucalcbol_Click()
Unload Frmboleta
Unload FrmCabezaBol
wTipoDoc = True
Load FrmCabezaBol
FrmCabezaBol.Show
FrmCabezaBol.ZOrder 0
End Sub

Private Sub mnucalcsegvida_Click()
NameForm = "SEGURODEVIDA"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnucarga_Click()
Load Frmcarga
Frmcarga.Show
Frmcarga.ZOrder 0
End Sub

Private Sub mnucargapla_Click()
'Load frmImport
'frmImport.Show
'frmImport.ZOrder 0
End Sub

Private Sub mnucargo_Click()
    FrmCargo.Show
    FrmCargo.ZOrder 0
End Sub

Private Sub mnucasca_Click()
Me.Arrange 0
End Sub

Private Sub mnucencos_Click()
    FrmCen_Costo.Show
End Sub

Private Sub mnucerafp_Click()
NameForm = "FrmCertretecionApf"
Load FrmCertretecionApf
FrmCertretecionApf.Show
FrmCertretecionApf.ZOrder 0
End Sub

Private Sub mnucerqta_Click()
NameForm = "CERTIFICAQTA"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnucia_Click()
Load Frmgrdcia
Frmgrdcia.Show
Frmgrdcia.ZOrder 0
End Sub

Private Sub mnuconcal_Click()
Load FrmDeduccion
FrmDeduccion.Show
FrmDeduccion.ZOrder 0
End Sub

Private Sub mnudeduccion_Click()
Load FrmDeduccion
FrmDeduccion.Show
FrmDeduccion.ZOrder 0
End Sub

Private Sub mnuConceptoRemu_Click()
    FrmRemunerativos.Show
    FrmRemunerativos.ZOrder 0
End Sub

Private Sub mnuconceptos_Click()
frmconceptos.Show
frmconceptos.ZOrder 0
End Sub

Private Sub MnuConder_Click()
Load FrmConDerechoHab
FrmConDerechoHab.Show
FrmConDerechoHab.ZOrder 0
End Sub

Private Sub mnuconremun_Click()
Load FrmRemuneraciones
FrmRemuneraciones.Show
FrmRemuneraciones.ZOrder 0
End Sub

Private Sub mnuconsultareo_Click()
Load Frmconsultareodia
Frmconsultareodia.Show
Frmconsultareodia.ZOrder 0
End Sub

Private Sub mnucont_Click()
Load frmcontratos
frmcontratos.Show
frmcontratos.ZOrder 0
End Sub

Private Sub mnucts_Click()
Load FrmSeteoCts
FrmSeteoCts.Show
FrmSeteoCts.ZOrder 0
End Sub

Private Sub MnuCtsFalta_Click()
Load FrmCtsFalta
FrmCtsFalta.Show
FrmCtsFalta.ZOrder 0
End Sub

Private Sub mnucudiv_Click()
NameForm = "CUADROIV"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnudeducapor_Click()
NameForm = "DEDUCAPOR"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnudepcts_Click()
Load FrmCts
FrmCts.Show
FrmCts.ZOrder 0
End Sub
Private Sub mnudepo_Click()
frmdatadepo.Show
End Sub

Private Sub mnudetquinta_Click()
NameForm = "DETALLEQUINTA"
Load Frmdetalle
Frmdetalle.Show
Frmdetalle.ZOrder 0
End Sub

Private Sub MnuDiasNoTrab_Click()
Load FrmDiasNoTrab
FrmDiasNoTrab.Show
FrmDiasNoTrab.ZOrder 0
End Sub

Private Sub MnuDiasSub_Click()
Load FrmDiasSub
FrmDiasSub.Show
FrmDiasSub.ZOrder 0
End Sub

Private Sub mnueps_Click()
    FrmEPS.Show
    FrmEPS.ZOrder 0
End Sub

Private Sub mnuexportar_Click()
    FrmExpotar.Show
End Sub

Private Sub mnufacsenati_Click()
Load Frmfactsenati
Frmfactsenati.Show
Frmfactsenati.ZOrder 0
End Sub

Private Sub mnufactor_Click()
Load FrmFactCalculo
FrmFactCalculo.Show
FrmFactCalculo.ZOrder 0
End Sub

Private Sub mnuforremu_Click()
Load FrmFormulasIng
FrmFormulasIng.Show
FrmFormulasIng.ZOrder 0
End Sub

Private Sub mnugenafpnet_Click()
Load FrmGenAPFNET
FrmGenAPFNET.Show
FrmGenAPFNET.ZOrder 0
End Sub

Private Sub mnuhelpcont_Click()
Call ShellExecute(hWnd, "Open", "c:\plared00\help\Help_Planilla.chm", "", "", vbNormalFocus)
End Sub

Private Sub mnuhor_Click()
Me.Arrange 1
End Sub

Private Sub mnuhoraper_Click()
Load Frmhorasper
Frmhorasper.Show
Frmhorasper.ZOrder 0
End Sub

Private Sub mnuhorasextras_Click()
    FrmHorasExtras.Show
    FrmHorasExtras.ZOrder 0
End Sub

Private Sub mnuhorcia_Click()
Load Frmverhoras
Frmverhoras.Show
Frmverhoras.ZOrder 0
End Sub

Private Sub mnuhorpla_Click()
Load FrmHoras
FrmHoras.Show
FrmHoras.ZOrder 0
End Sub

Private Sub mnuiniciosem_Click()
Load Frminicosemana
Frminicosemana.Show
Frminicosemana.ZOrder 0
End Sub

Private Sub mnulistquinta_Click()
NameForm = "LISTAQUINTA"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnulegal_Click()
frmplalegal.Show
End Sub

Private Sub mnumae_Click()
Load frmtmpmae
frmtmpmae.Show
frmtmpmae.ZOrder 0
End Sub

Private Sub mnuobras_Click()
wobra = True
Load Frmgrdobra
Frmgrdobra.Show
Frmgrdobra.ZOrder 0
End Sub

Private Sub mnuorgico_Click()
Me.Arrange 3
End Sub

Private Sub mnuperpago_Click()
Load Frmmodopago
Frmmodopago.Show
Frmmodopago.ZOrder 0
End Sub

Private Sub Mnupersona_Click()
Load FrmGrdPersonal
FrmGrdPersonal.Show
FrmGrdPersonal.ZOrder 0
End Sub

Private Sub MnuPlaElec_Click()
Load FrmExpPe
FrmExpPe.Show
FrmExpPe.ZOrder 0
End Sub

Private Sub mnuplamas_plahistorico_Click()
    Frm_Plamas_Plahistorico.Show
    Frm_Plamas_Plahistorico.ZOrder 0
End Sub

Private Sub mnupresta_Click()
Load Frmgrdctacte
Frmgrdctacte.Show
Frmgrdctacte.ZOrder 0
End Sub

Private Sub mnuprintbol_Click()
Unload Frmprintboleta
Frmprintboleta.wmBolQuin = "B"
Frmprintboleta.Show
Frmprintboleta.ZOrder 0
'Unload frmprintcrystal
'frmprintcrystal.wmBolQuin = "B"
'frmprintcrystal.Show
'frmprintcrystal.ZOrder 0
End Sub

Private Sub mnuprintquinc_Click()
Unload Frmprintboleta
Frmprintboleta.wmBolQuin = "Q"
Frmprintboleta.Show
Frmprintboleta.ZOrder 0
End Sub

Private Sub mnuprom_Click()
Load Frmpromedio
Frmpromedio.Show
Frmpromedio.ZOrder 0
End Sub

Private Sub mnuprovac_Click()
Unload Frmprovision
Frmprovision.Provisiones ("V")
Frmprovision.Show
Frmprovision.ZOrder 0
End Sub

Private Sub mnuprovgrati_Click()
Unload Frmprovision
Frmprovision.Provisiones ("G")
Frmprovision.Show
Frmprovision.ZOrder 0
End Sub

Private Sub mnuquincena_Click()
Unload Frmboleta
Unload FrmCabezaBol
wTipoDoc = False
Load FrmCabezaBol
FrmCabezaBol.Show
FrmCabezaBol.ZOrder 0
End Sub

Private Sub mnuregfer_Click()
Load Frmferiados
Frmferiados.Show
Frmferiados.ZOrder 0
End Sub

Private Sub mnuremu_cts_Click()
    FrmRemuneracion_CTS.Show
    FrmRemuneracion_CTS.ZOrder 0
End Sub

Private Sub mnuremun_Click()
NameForm = "REMUNERA"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnuremunrasunat_Click()
NameForm = "REMUNERACIONES"
Load Frmbarraprogress
Frmbarraprogress.Show
Frmbarraprogress.ZOrder 0
End Sub

Private Sub mnurestaura_Click()
Unload Frmbackup
Frmbackup.Opc = "R"
Load Frmbackup
Frmbackup.Show
Frmbackup.ZOrder 0
End Sub

Private Sub mnuresumen_Click()
NameForm = "RESUMEN"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnusctr_Click()
Load FrmSCTR
FrmSCTR.Show
FrmSCTR.ZOrder 0
End Sub

Private Sub mnuseg_Click()
NameForm = "SEGURO"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub mnuselecc_Click()
Load Frmseleccion
Frmseleccion.Show
Frmseleccion.ZOrder 0
End Sub

Private Sub mnuseteobol_Click()
Load Frmseteoimpresion
Frmseteoimpresion.Show
Frmseteoimpresion.ZOrder 0
End Sub

Private Sub mnusub_Click()
Load Frmgrdsubsidio
Frmgrdsubsidio.Show
Frmgrdsubsidio.ZOrder 0
End Sub

Private Sub mnusunattrab_Click()
NameForm = "TRABAJADORES"
Load Frmbarraprogress
Frmbarraprogress.Show
Frmbarraprogress.ZOrder 0
End Sub

Private Sub mnutareo_Click()
Load Frmgrdtareo
Frmgrdtareo.Show
Frmgrdtareo.ZOrder 0
End Sub

Private Sub mnutasascrt_Click()
Load Frmcalculoscrs
Frmcalculoscrs.Show
Frmcalculoscrs.ZOrder 0
End Sub

Private Sub mnutc_Click()
Load Frmtipcamb
Frmtipcamb.Show
Frmtipcamb.ZOrder 0
End Sub

Private Sub mnuuit_Click()
Load FrmUit
FrmUit.Show
FrmUit.ZOrder 0
End Sub

Private Sub mnuusers_Click()
Load Frmusers
Frmusers.Show
Frmusers.ZOrder 0
End Sub

Private Sub mnuvacdev_Click()
Unload Frmprovision
Frmprovision.Provisiones ("D")
Frmprovision.Show
Frmprovision.ZOrder 0
End Sub

Private Sub mnuver_Click()
Me.Arrange 2
End Sub

Private Sub mnuvf_Click()
NameForm = "CUADROIVF"
Load Frmmes
Frmmes.Show
Frmmes.ZOrder 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.index
Case Is = 10
     If DoEvents > 1 Then
        On Error GoTo Salir
        If Screen.ActiveForm.Name = "MDIplared" Then
           MsgBox "Debe Cerrar Todos los Formularios,Tal vez estén minimizados", vbInformation
        Else
           If MsgBox("Desea Cerrar el Formulario", vbYesNo + vbQuestion) = vbYes Then
              Unload MDIplared.ActiveForm
           End If
        End If
     Else
        If MsgBox("Desea Salir Del Sistema", vbYesNo + vbQuestion) = vbYes Then Unload MDIplared
     End If
Case Is = 1
     If DoEvents > 1 Then
        Select Case MDIplared.ActiveForm.Name
        Case Is = "aFrmDH"
             aFrmDH.NuevoTrab
        Case Is = "FrmConDerechoHab"
             aFrmDH.NuevoTrab
        Case Is = "FrmDiasSub"
             FrmDiasSub.Nuevo
        Case Is = "FrmDiasNoTrab"
            FrmDiasNoTrab.Nuevo
        Case Is = "FrmGrdPersonal"
             Frmpersona.Nuevo_Personal (True)
        Case Is = "Frmpersona"
             Frmpersona.Nuevo_Personal (True)
        Case Is = "Frmplacte"
             Frmplacte.Nuevo_Prestamo
        Case Is = "Frmgrdctacte"
             Frmplacte.Nuevo_Prestamo
        Case Is = "Frmboleta"
             Frmboleta.Limpia_Boleta
        Case Is = "Frmcia"
             Frmcia.Nueva_cia
        Case Is = "Frmgrdcia"
             Frmcia.Nueva_cia
        Case Is = "Frmusers"
             Frmusers.Nuevo_User
        Case Is = "Frmgrdobra"
             If wobra = True Then Frmobras.Nueva_Obra (Frmgrdobra.Cmbcia.Text)
        Case Is = "Frmobras"
             Frmobras.Nueva_Obra (Frmobras.LblCia.Caption)
        Case Is = "Frmtareo"
             Frmtareo.Nuevo_Tareo
        Case Is = "Frmgrdtareo"
             Frmtareo.Nuevo_Tareo
        Case Is = "Frmgrdsubsidio"
             Frmsubsidios.Limpiar
        Case Is = "FrmCtsFalta"
             FrmCtsFaltaEdit.Nuevo
        End Select
     End If
Case Is = 2  'GRABAR
     DoEvents
     If DoEvents > 1 Then
        Select Case MDIplared.ActiveForm.Name
                Case Is = "aFrmDH"
                    aFrmDH.Grabar
               Case Is = "FrmDiasNoTrab"
                    FrmDiasNoTrab.Grabar
               Case Is = "FrmDiasSub"
                    FrmDiasSub.Grabar
               Case Is = "FrmGrdPersonal"
                    If FrmGrdPersonal.FrameReport.Visible = True Then FrmGrdPersonal.Grabar_Report
               Case Is = "Frmpersona" 'GRABAR NUEVO EMPLEADO
                    Frmpersona.Grabar_Persona
                Case Is = "frmconceptos"
                    frmconceptos.GrabaConceptos
               Case Is = "FrmAfp"
                    FrmAfp.GrabarAfp
               Case Is = "FrmSCTR"
                    FrmSCTR.GrabarSCTR
               Case Is = "FrmDeduccion"
                    FrmDeduccion.Graba_Concepto
               Case Is = "Frmcalculoscrs"
                    Frmcalculoscrs.Graba_Seguro
               Case Is = "FrmRemuneraciones" 'GRABA REMUNERACIONES
                    FrmRemuneraciones.Graba_Remunera
               Case Is = "FrmUit"
                    FrmUit.Graba_Uit
               Case Is = "FrmFactCalculo"
                    FrmFactCalculo.Graba_FactorCalculo
               Case Is = "Frmpromedio"
                    Frmpromedio.Graba_Promedio
               Case Is = "FrmFormulasIng"
                    FrmFormulasIng.Graba_FromulaIng
               Case Is = "Frmplacte"
                    Frmplacte.Graba_Prestamo
               Case Is = "Frmcia"
                    Frmcia.Graba_Cia
               Case Is = "Frmboleta"
                    Frmboleta.Grabar_Boleta
               Case Is = "Frmmodopago"
                    Frmmodopago.GrabarPerPago
               Case Is = "Frmhorasper"
                    Frmhorasper.Grabar_PerHoras
               Case Is = "frmtmpmae2"
                    'frmtmpmae2.Grabar
               Case Is = "Frmtipcamb"
                    Frmtipcamb.Graba_TIpCamb
               Case Is = "Frmusers"
                    Frmusers.Grabar_users
               Case Is = "Frmobras"
                    Frmobras.Graba_Obra
               Case Is = "Frmtareo"
                    Frmtareo.Grabar_Tareo
               Case Is = "FrmHoras"
                    FrmHoras.GrabarHorasPla
               Case Is = "FrmAfectos"
                    FrmAfectos.Grabar_Afectos
               Case Is = "Frmverhoras"
                    Frmverhoras.Grabar_VerHoras
               Case Is = "Frminicosemana"
                    Frminicosemana.Grabar_InicioSem
               Case Is = "Frmseteoimpresion"
                    Frmseteoimpresion.Grabar_Seteo_Print
               Case Is = "Frmfactsenati"
                    Frmfactsenati.Grabar_Senati
               Case Is = "Frmsubsidios"
                    Frmsubsidios.Grabar_Subsidio
               Case Is = "FrmSeteoCts"
                    FrmSeteoCts.Grabar_Seteo_Cts
               Case Is = "Frmferiados"
                    Frmferiados.Grabar_Feriados
               Case Is = "Frmseleccion"
                    Frmseleccion.Grabar_Report_Seleccion
               Case Is = "frmdesccta"
                    frmdesccta.guardar1
               Case Is = "FrmMContable"
                    FrmMContable.Graba_Informacion_Ingresada
               Case Is = "FrmMContableFijo"
                    FrmMContableFijo.Grabar_Informacion_Nueva
               Case Is = "FrmMContAsoc"
                    FrmMContAsoc.Grabar_Informacion_Nueva
                Case Is = "FrmCtaProvision"
                    FrmCtaProvision.Graba_Informacion_Ingresada
                Case Is = "FrmCtsFaltaEdit"
                    FrmCtsFaltaEdit.Grabar
        End Select
     End If
Case Is = 3     'ELIMINAR
     If DoEvents > 1 Then
        Select Case MDIplared.ActiveForm.Name
        Case Is = "FrmDiasNoTrab"
             FrmDiasNoTrab.Elimimar
        Case Is = "FrmDiasSub"
             FrmDiasSub.Elimimar
        Case Is = "Frmboleta"
             Frmboleta.Elimina_Boleta
        Case Is = "Frmusers"
             Frmusers.Eliminar_Usuario
        Case Is = "Frmobras"
             Frmobras.Elimina_Obra
        Case Is = "Frmtareo"
             Frmtareo.Elimina_Tareo
        Case Is = "Frmgrdsubsidio"
             Frmgrdsubsidio.Elimina_Subsidio
        Case Is = "FrmCts"
             FrmCts.Elimina_Cts
        Case Is = "Frmprovision"
             Select Case Frmprovision.Lbltipo
                    Case Is = "V": Frmprovision.Elimina_Prov_Vaca
                    Case Is = "G": Frmprovision.Elimina_Prov_Grati
             End Select
        Case Is = "FrmRemuneraciones"
              FrmRemuneraciones.Eliminar
        Case Is = "FrmCtsFalta"
             FrmCtsFalta.Eliminar
        End Select
     End If
Case Is = 5  'BUSCAR
     If DoEvents > 1 Then
        Select Case MDIplared.ActiveForm.Name
        Case Is = "Frmplacte"
             Load Frmgrdpla
             Frmgrdpla.Visible = True
             Frmgrdpla.ZOrder 0
        Case Is = "Frmgrdctacte"
             Load Frmgrdpla
             Frmgrdpla.Visible = True
             Frmgrdpla.ZOrder 0
        Case Is = "Frmboleta"
             If Frmboleta.Txtcodpla.Enabled = True Then
                Load Frmgrdpla
                Frmgrdpla.Visible = True
                Frmgrdpla.ZOrder 0
             End If
        Case Is = "Frmpersona"
             If wbus = "OB" Then
                wobra = False
                Load Frmgrdobra
                Frmgrdobra.Show
                Frmgrdobra.ZOrder 0
             End If
        Case Is = "FrmCabezaBol"
             If wbus = "OB" Then
                wobra = False
                Load Frmgrdobra
                Frmgrdobra.Show
                Frmgrdobra.ZOrder 0
             End If
        Case Is = "Frmtareo"
             If wbus = "PL" Then
                Load Frmgrdpla
                Frmgrdpla.Visible = True
                Frmgrdpla.ZOrder 0
             End If
             If wbus = "OB" Then
                wobra = False
                Load Frmgrdobra
                Frmgrdobra.Show
                Frmgrdobra.ZOrder 0
             End If
        Case Is = "Frmgrdtareo"
             If wbus = "PL" Then
                Load Frmgrdpla
                Frmgrdpla.Visible = True
                Frmgrdpla.ZOrder 0
             End If
             If wbus = "OB" Then
                wobra = False
                Load Frmgrdobra
                Frmgrdobra.Show
                Frmgrdobra.ZOrder 0
             End If
        Case Is = "Frmdetalle"
                Load Frmgrdpla
                Frmgrdpla.Visible = True
                Frmgrdpla.ZOrder 0
        End Select
     End If
Case Is = 7
     If DoEvents > 1 Then
     
        Select Case MDIplared.ActiveForm.Name
        Case Is = "frmprintcrystal"
             frmprintcrystal.Imprime_Boletas
        Case Is = "Frmprintboleta"
             Frmprintboleta.Imprime_Boletas
        Case Is = "FrmGrdPersonal"
             If FrmGrdPersonal.FrameReport.Visible = False Then FrmGrdPersonal.Reporte_Personal
        Case Is = "FrmMGenera"
                FrmMGenera.Proceso_Reporte_Detallado
        Case Is = "FrmRptSeleccion"
                FrmRptSeleccion.Procesar_Reporte
        Case Is = "FrmRptseleccion2"
                FrmRptseleccion2.Proceso_Reporte_Dos
        Case Is = "FrmRptAportes"
                FrmRptAportes.Proceso_Ejecuta_Reporte
        Case Is = "FrmRptRetenciones"
                FrmRptRetenciones.Proceso_Reporte_Central
        Case Is = "FrmRptPromedios"
                FrmRptPromedios.Proceso_Central_Reporte_Promedios
        Case Is = "Frmgrdctacte"
             If Frmgrdctacte.Framemovi.Visible = True Then
                Frmgrdctacte.Imprime_Movimientos
             Else
                Frmgrdctacte.Imprime_Saldos
             End If
        Case Else
             wPrintFile = ""
             Load Formimpri
             Formimpri.Show
             Formimpri.ZOrder 0
        End Select
     Else
        Load Formimpri
        Formimpri.Show
        Formimpri.ZOrder 0
     End If
Case Is = 8
     If DoEvents > 1 Then
        Select Case MDIplared.ActiveForm.Name
        Case Is = "FrmRepafp"
             FrmRepafp.Procesa_RepAfp
        Case Is = "FrmGenAPFNET"
            FrmGenAPFNET.Procesa_RepAfpNet
        Case Is = "Frmmes"
             Frmmes.Procesar
        Case Is = "Frmdetalle"
             Frmdetalle.Procesar_Detalle
        Case Is = "Frmafpaunal"
             Frmafpaunal.Procesa_Afp_Anual
        Case Is = "frmdatadepo"
            frmdatadepo.Procesar
        Case Is = "FrmMGenera"
                FrmMGenera.Proceso_General
        Case Is = "FrmGenerarDBF"
                FrmGenerarDBF.Generar_Archivo_Dbf
        End Select
     End If
Case Is = 16
    If DoEvents > 1 Then
       If Screen.ActiveForm.Name = "MDIplared" Then
          MsgBox "Debe Cerrar Todos los Formularios, Tal vez estén minimizados", vbInformation, "Verifique"
       Else
         MsgBox "Debe Cerrar Todos los Formularios, Para poder Cambiar de Compañia", vbInformation, "Verifique"
       End If
    Else
       MDIplared.Enabled = False
       Load FrmSelCia
       FrmSelCia.Show 1
       Call Setea_StatusBar
    End If
     
End Select
Exit Sub
Salir: Unload MDIplared
End Sub
Public Sub Activa_Menu()
If StrConv(wuser, 1) = wAdmin Then Exit Sub
Dim MiObjeto
For Each MiObjeto In MDIplared
    If TypeOf MiObjeto Is Menu Then
       If MiObjeto.Caption <> "-" Then
          Sql$ = "select * from users_menu where sistema='" & wCodSystem & "' and status<>'*' and cia='" & wcia & "' and name_user='" & wuser & "' and name_menu='" & MiObjeto.Name & "'"
          cn.CursorLocation = adUseClient
          Set rs = New ADODB.Recordset
          Set rs = cn.Execute(Sql$, 64)
          If rs.RecordCount > 0 Then
             MiObjeto.Enabled = True
          Else
             MiObjeto.Enabled = False
          End If
          If MiObjeto.Name = "mnufacsenati" Then
             If wSenati = "S" Then MiObjeto.Visible = True Else MiObjeto.Visible = False
          End If
          m = m + 1
       End If
    End If
Next
If rs.State = 1 Then rs.Close
End Sub

Public Sub Carga_DerechoHabientes()
Dim Sql As String
Sql = "SELECT * FROM TMPPDTDERE2 order by PLACOD"
Dim Rq As ADODB.Recordset
Dim xTrab As String
xTrab = ""
Dim VCodAuxi As String
Dim NumAuxi As String
Dim xNombres1 As String
Dim xNombres2 As String
xNombres1 = ""
xNombres2 = ""
Dim xCodTrab As String
xCodTrab = ""
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    cn.BeginTrans
    cn.Execute "DELETE from pladerechohab where placod in (select placod from TMPPDTDERE2) and status<>'*' ", 64
    Do While Not Rq.EOF
            If Trim(xTrab) <> Trim(Rq!PlaCod) Then
                VCodAuxi = "": NumAuxi = ""
                Dim Rp As ADODB.Recordset
                Sql = "select * from planillas where placod='" & Trim(Rq!PlaCod) & "' and status<>'*' AND fcese IS NULL"
                If Not fAbrRst(Rp, Sql) Then
                    Sql = "select * from planillas where placod='" & Trim(Rq!PlaCod) & "' and status<>'*'"
                    If Not fAbrRst(Rp, Sql) Then
                        Debug.Print "FALTA TRABAJADOR " & Trim(Rq!PlaCod)
                        xCodTrab = ""
                    Else
                        xCodTrab = Trim(Rp!PlaCod & "")
                        VCodAuxi = Trim(Rp!codauxinterno & "")
                        'Stop
                    End If
                Else
                    xCodTrab = Trim(Rp!PlaCod & "")
                    VCodAuxi = Trim(Rp!codauxinterno & "")
                End If
                
            End If
'            Dim xBlanco As Integer
'            xBlanco = InStr(1, Rq!nombre, " ")
'            If xBlanco > 0 Then
'                xNombres1 = Mid(Rq!nombre, 1, xBlanco - 1)
'                xNombres2 = Mid(Rq!nombre, xBlanco + 1, Len(Rq!nombre) - Len(xNombres1))
'            Else
'                xNombres1 = Trim(Rq!nombre & "")
'                xNombres2 = ""
'            End If
            xNombres1 = Left(Trim(Rq!nom_1 & ""), 20)
            xNombres2 = Left(Trim(Rq!nom_2 & ""), 20)
            
'            Dim xTipDocOrder As String
'            Dim xNroDocOrder As String
'            Select Case Trim(Rq!TipDocder)
'            Case "1":
'                xTipDocOrder = "01" 'dni
'                xNroDocOrder = Trim(Rq!NumDocder1)
'            Case "11"
'                xTipDocOrder = "08" 'partida nac
'                xNroDocOrder = Trim(Rq!NumDocder1)
'            Case Else
'                xTipDocOrder = ""
'                xNroDocOrder = ""
'            End Select
            Dim xFecNac As String
            xFecNac = Format(Rq!fec_nac, "dd/mm/yyyy")
            xFecNac = Format(xFecNac, "mm/dd/yyyy")
            Dim xTipDoc As String
            Select Case Trim(Rq!TIP_DOC & "")
            Case "11": xTipDoc = "08" 'PARTIDA
            Case "07": xTipDoc = "05" 'PASAPORTE
            Case Else
                xTipDoc = Trim(Rq!TIP_DOC & "")
            End Select
            If Trim(xCodTrab) <> "" Then
                Sql = "insert into pladerechohab values('" & wcia & "','" & xCodTrab & "','" & VCodAuxi
                Sql = Sql & "','" & Left(Trim(Rq!ap_pat), 20) & "','" & Left(Trim(Rq!ap_mat), 20) & "','" & Apostrofe(Trim(xNombres1)) & "','" & Apostrofe(Trim(xNombres2))
                Sql = Sql & "','" & Trim(xTipDoc) & "','" & IIf(Trim(xTipDoc & "") <> "", Trim(Rq!num_doc), "") & "','" & xFecNac & "','" & Format(Rq!vinculo, "00") & "','','" & Rq!sexo & "'"
                Sql = Sql & ",'1','','N','','S','',getdate(),'','-1','','','-1','','-1','','','','','','',null,null,NULL)"
                cn.Execute Sql, 64
            End If
            xTrab = Trim(Rq!PlaCod)
            
        Rq.MoveNext
    Loop
    cn.CommitTrans
End If
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
MsgBox "Termino la carga", vbInformation

End Sub

Public Sub Actualiza_DatosTrabajador()
Dim Sql As String

Sql = "SELECT P.PLACOD,P.nro_doc,T.DNI ,p.ap_pat,p.ap_pat,p.ap_mat,t.ape_mat,p.nom_1,t.nombres"
Sql = Sql & " ,p.fnacimiento,t.FEC_NAC,p.niveleducativo,t.NIVEL_EDU,p.tipo_contrato,t.TIP_CONTRATO"
Sql = Sql & " ,p.trab_situacion_especial,t.SITUACION_ESPECIAL"
Sql = Sql & " ,p.trab_reg_alternativo,t.[14]"
Sql = Sql & " ,p.trab_jornada_trab_max,t.[15]"
Sql = Sql & " ,p.trab_hor_nocturno,t.[16]"
Sql = Sql & " ,p.tvia ,t.COD_VIA"
Sql = Sql & " ,p.nomvia,t.NOM_VIA"
Sql = Sql & " ,p.tzona,t.COD_ZONA"
Sql = Sql & " ,p.nomzona,t.NOM_ZONA"
Sql = Sql & " ,P.ubigeo,T.COD_UBIGEO"
Sql = Sql & " ,T.NUMERO,T.VIA "
Sql = Sql & " FROM PLANILLAS P,TMPMONTORO T"
Sql = Sql & " Where P.PLACOD = T.codiGO order by placod"


Dim Rq As ADODB.Recordset
Dim xTrab As String
xTrab = ""
Dim VCodAuxi As String
Dim NumAuxi As String
Dim xNombres1 As String
Dim xNombres2 As String
xNombres1 = ""
xNombres2 = ""
Dim xCodTrab As String
xCodTrab = ""
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    cn.BeginTrans
    Do While Not Rq.EOF
            
            xCodTrab = Trim(Rq!PlaCod & "")
            
            Dim xBlanco As Integer
            xBlanco = InStr(1, Rq!nombres, " ")
            If xBlanco > 0 Then
                xNombres1 = Mid(Rq!nombres, 1, xBlanco - 1)
                xNombres2 = Mid(Rq!nombres, xBlanco + 1, Len(Rq!nombres) - Len(xNombres1))
            Else
                xNombres1 = Trim(Rq!nombres & "")
                xNombres2 = ""
            End If
            
'            Dim xTipDocOrder As String
'            Dim xNroDocOrder As String
'            Select Case Trim(Rq!TipDocder)
'            Case "1":
'                xTipDocOrder = "01" 'dni
'                xNroDocOrder = Trim(Rq!NumDocder1)
'            Case "11"
'                xTipDocOrder = "08" 'partida nac
'                xNroDocOrder = Trim(Rq!NumDocder1)
'            Case Else
'                xTipDocOrder = ""
'                xNroDocOrder = ""
'            End Select
'            Dim xFecNac As String
'            xFecNac = Format(Rq!FechaNacimi, "dd/mm/yyyy")
'            xFecNac = Format(xFecNac, "mm/dd/yyyy")
'            If Trim(xCodTrab) <> "" Then
'                Sql = "insert into pladerechohab values('" & wcia & "','" & xCodTrab & "','" & VCodAuxi
'                Sql = Sql & "','" & Trim(Rq!ApPaterno) & "','" & Trim(Rq!ApMaterno) & "','" & Apostrofe(Trim(xNombres1)) & "','" & Apostrofe(Trim(xNombres2))
'                Sql = Sql & "','" & xTipDocOrder & "','" & xNroDocOrder & "','" & xFecNac & "','" & Format(Rq!vinculo, "00") & "','','" & IIf(Rq!sexo = 1, "M", "F") & "'"
'                Sql = Sql & ",'1','','N','','S','',getdate(),'','-1','','','-1','','-1','','','','','','',null,null,NULL)"
'                cn.Execute Sql, 64
'            End If
            Sql = "update planillas set "
            Sql = Sql & " nro_doc='" & Trim(Rq!dni & "") & "'"
            Sql = Sql & " ,ap_pat='" & Trim(Rq!ap_pat & "") & "'"
            Sql = Sql & " ,ap_mat='" & Trim(Rq!ape_mat & "") & "'"
            Sql = Sql & " ,nom_1='" & xNombres1 & "'"
            Sql = Sql & " ,nom_2='" & xNombres2 & "'"
            Sql = Sql & " ,fnacimiento='" & Format(Rq!fec_nac, "mm/dd/yyyy") & "'"
            Sql = Sql & " ,niveleducativo='" & Trim(Rq!NIVEL_EDU & "") & "'"
            Sql = Sql & " ,tipo_contrato='" & Trim(Rq!TIP_CONTRATO & "") & "'"
            Sql = Sql & " ,trab_siTuacion_especial='" & Trim(Rq!SITUACION_ESPECIAL & "") & "'"
            Sql = Sql & " ,trab_reg_alternativo='" & Trim(Rq.Fields("14") & "") & "'"
            Sql = Sql & " ,trab_jornada_trab_max='" & Trim(Rq.Fields("15") & "") & "'"
            Sql = Sql & " ,trab_hor_nocturno='" & Trim(Rq.Fields("16") & "") & "'"
            Sql = Sql & " ,tvia ='" & Trim(Rq!cod_via & "") & "'"
            Sql = Sql & " ,nomvia='" & Trim(Rq!NOM_VIA & "") & "'"
            
            Sql = Sql & " ,nrokmmza='" & Trim(Rq!numero & "") & "'"
            Sql = Sql & " ,intdptolote='" & Trim(Rq!via & "") & "'"
            
            Sql = Sql & " ,tzona='" & Trim(Rq!cod_zona & "") & "'"
            Sql = Sql & " ,nomzona='" & Trim(Rq!NOM_ZONA & "") & "'"
            Sql = Sql & " ,ubigeo='" & Trim(Rq!COD_UBIGEO & "") & "'"
            Sql = Sql & " where placod='" & xCodTrab & "'"
            cn.Execute Sql, 64
             
        Rq.MoveNext
    Loop
    cn.CommitTrans
End If
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
MsgBox "Termino la carga", vbInformation

End Sub


Public Sub Actualiza_DerechoHabientes()
Dim Sql As String
Sql = "SELECT * FROM TMPPDTDERE2 order by placod"
Dim Rq As ADODB.Recordset
Dim xTrab As String
xTrab = ""
Dim VCodAuxi As String
Dim NumAuxi As String
Dim xNombres1 As String
Dim xNombres2 As String
xNombres1 = ""
xNombres2 = ""
Dim xCodTrab As String
xCodTrab = ""
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    cn.BeginTrans
    'cn.Execute "delete from pladerechohab", 64
    Do While Not Rq.EOF
            Dim xFecNac As String
            xFecNac = Format(Rq!fec_nac, "dd/mm/yyyy")
            xFecNac = Format(xFecNac, "mm/dd/yyyy")
            Sql = "update pladerechohab set ap_pat='" & Trim(Rq!ap_pat) & "',ap_mat='" & Trim(Rq!ap_mat)
            Sql = Sql & "',nom_1='" & Trim(Rq!nom_1) & "',nom_2='" & Trim(Rq!nom_2) & "',fec_nac='" & xFecNac & "'"
            Sql = Sql & ",sexo='" & Trim(Rq!sexo) & "',cod_doc='" & Trim(Rq!TIP_DOC) & "',numero='" & Trim(Rq!num_doc) & "',codvinculo='" & Rq!vinculo & "'"
            Sql = Sql & " where placod='" & Trim(Rq!PlaCod) & "' and status<>'*'"
            cn.Execute Sql, 64
        Rq.MoveNext
    Loop
    cn.CommitTrans
End If
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
MsgBox "Termino la carga", vbInformation

End Sub

